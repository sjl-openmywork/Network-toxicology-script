import os
import sys
from rdkit import Chem
from rdkit.Chem import (
    Descriptors, Lipinski, Crippen, rdMolDescriptors, AllChem
)
from rdkit import DataStructs
import pandas as pd
import re
from collections import Counter

# ==================== 新增依赖说明 ====================
# 本版本支持 mol2 格式输出，需要额外安装 OpenBabel 的 Python 绑定
# 推荐安装方式：
#   conda install -c conda-forge openbabel
# 或（Windows 用户推荐）：
#   pip install openbabel-wheel
# 如果未安装 OpenBabel，选择 mol2 时会自动回退到 sdf 并提示
# ====================================================

# 常见官能团 SMARTS 定义
FG_PATTERNS = {
    '羧酸': 'C(=O)[OX2H1]',
    '酚羟基': 'c[OH]',
    '胺': '[NX3;H2,H1,H0][CX4]',
    '酰胺': 'C(=O)[NX3]',
    '酯': '[CX3](=O)[OX2][CX4]',
    '醚': '[OD2]([CX4])[CX4]',
    '醛': '[CX3H1](=O)[CX3]',
    '酮': '[CX3](=O)[CX3;!R]',
    '硝基': '[$([NX3](=O)=O)]',
    '卤素': '[F,Cl,Br,I]'
}

SUPPORTED_EXT = ['.xlsx', '.sdf', '.mol', '.mol2', '.pdb']

# 英文文件夹名 -> 中文显示映射
FOLDER_MAP = {
    "Molecular_Info": "分子信息",
    "Format_Conversion": "格式转换",
    "Similarity_Comparison": "相似性比较"
}

# 支持的输出格式（mol2 依赖 OpenBabel）
SUPPORTED_OUTPUT_FORMATS = ['sdf', 'mol', 'pdb', 'mol2']

def safe_filename(name):
    return re.sub(r'[<>:"/\\|?*]', '_', str(name))

# 解析文件（同前）
def parse_file(file_path):
    print(f"  正在解析: {os.path.basename(file_path)}")
    ext = os.path.splitext(file_path)[1].lower()
    source = os.path.basename(file_path)
    mols = []
    mol_names = []

    if ext == '.xlsx':
        df = pd.read_excel(file_path)
        if not {'names', 'smiles'}.issubset(df.columns):
            raise ValueError(f"{source} 必须包含 'names' 和 'smiles' 两列")
        for idx, (name, smi) in enumerate(zip(df['names'], df['smiles'])):
            mol = Chem.MolFromSmiles(str(smi))
            if mol is None:
                print(f"  警告：无效SMILES，跳过 -> {smi}")
                continue
            AllChem.Compute2DCoords(mol)
            mol_name = str(name) if pd.notna(name) else f"Mol_{idx+1}"
            mols.append(mol)
            mol_names.append(mol_name)

    elif ext == '.sdf':
        supplier = Chem.SDMolSupplier(file_path, sanitize=False)
        for idx, mol in enumerate(supplier):
            if mol is None:
                continue
            AllChem.Compute2DCoords(mol)
            mol_name = mol.GetProp("_Name") if mol.HasProp("_Name") else f"Mol_{idx+1}"
            mols.append(mol)
            mol_names.append(mol_name)

    elif ext == '.mol':
        mol = Chem.MolFromMolFile(file_path, sanitize=False)
        if mol is None:
            raise ValueError(f"无法读取 {source}")
        AllChem.Compute2DCoords(mol)
        mol_name = os.path.splitext(source)[0]
        mols.append(mol)
        mol_names.append(mol_name)

    elif ext == '.mol2':
        mol = Chem.MolFromMol2File(file_path)
        if mol is None:
            raise ValueError(f"无法读取 {source}")
        AllChem.Compute2DCoords(mol)
        mol_name = os.path.splitext(source)[0]
        mols.append(mol)
        mol_names.append(mol_name)

    elif ext == '.pdb':
        mol = Chem.MolFromPDBFile(file_path, sanitize=False)
        if mol is None:
            raise ValueError(f"无法读取 {source}")
        AllChem.Compute2DCoords(mol)
        mol_name = os.path.splitext(source)[0]
        mols.append(mol)
        mol_names.append(mol_name)

    return mols, mol_names, source

# 属性计算（同前）
def compute_properties(mols, names, sources, add_source_column):
    print("正在计算分子属性...")
    props = []
    prop_keys = ['名称', 'SMILES', '分子量', 'LogP', 'TPSA', '氢键供体', '氢键受体',
                 '可旋转键', '环数', '重原子数', '常见官能团']

    for i, mol in enumerate(mols):
        canon_smiles = Chem.MolToSmiles(mol, canonical=True)

        present_fgs = []
        for group, smarts in FG_PATTERNS.items():
            patt = Chem.MolFromSmarts(smarts)
            if patt and mol.HasSubstructMatch(patt):
                present_fgs.append(group)

        data = {
            '名称': names[i],
            'SMILES': canon_smiles,
            '分子量': round(Descriptors.MolWt(mol), 2),
            'LogP': round(Crippen.MolLogP(mol), 2),
            'TPSA': round(rdMolDescriptors.CalcTPSA(mol), 2),
            '氢键供体': Lipinski.NumHDonors(mol),
            '氢键受体': Lipinski.NumHAcceptors(mol),
            '可旋转键': Lipinski.NumRotatableBonds(mol),
            '环数': rdMolDescriptors.CalcNumRings(mol),
            '重原子数': Lipinski.HeavyAtomCount(mol),
            '常见官能团': ', '.join(present_fgs) if present_fgs else '无'
        }
        if add_source_column:
            data['来源文件'] = sources[i]
        props.append(data)

    df = pd.DataFrame(props)
    order = ['来源文件'] + prop_keys if add_source_column else prop_keys
    return df[order]

# 3D处理函数（同前）
def generate_optimize_3d(mol, name):
    m = Chem.Mol(mol)
    try:
        Chem.SanitizeMol(m)
    except:
        print(f"  警告：{name} 初始结构有问题，尝试继续")

    has_3d = m.GetNumConformers() > 0 and m.GetConformer().Is3D()

    m = Chem.AddHs(m)

    if has_3d:
        print(f"  {name}: 检测到原始3D → 保留坐标，仅添加氢并优化")
    else:
        print(f"  {name}: 无原始3D → 使用ETKDGv3生成3D构象")
        params = AllChem.ETKDGv3()
        params.randomSeed = 42
        params.pruneRmsThresh = 0.5
        res = AllChem.EmbedMolecule(m, params)
        if res != 0:
            print(f"  警告：{name} 单次嵌入失败 → 尝试生成10个构象选取最佳")
            AllChem.EmbedMultipleConfs(m, numConfs=10, params=params)

    print(f"  {name}: 正在优化...", end=' ')
    try:
        if AllChem.MMFFHasAllMoleculeParams(m):
            res = AllChem.MMFFOptimizeMolecule(m)
            if res == 0:
                print("MMFF优化成功")
            else:
                AllChem.UFFOptimizeMolecule(m)
                print("MMFF失败 → UFF优化")
        else:
            AllChem.UFFOptimizeMolecule(m)
            print("UFF优化")
    except Exception as e:
        print(f"优化失败（将继续保存）：{e}")

    if m.GetNumConformers() == 0:
        print(f"  错误：{name} 最终无3D → 回退2D坐标")
        AllChem.Compute2DCoords(m)

    return m

# 格式转换（新增 mol2 支持，依赖 OpenBabel）
def convert_format(mols, names, target_format, output_dir):
    # 检查格式并处理 OpenBabel 依赖
    if target_format not in SUPPORTED_OUTPUT_FORMATS:
        print(f"不支持的格式 {target_format}，已回退到 sdf")
        target_format = 'sdf'

    use_openbabel = (target_format == 'mol2')
    pybel = None
    if use_openbabel:
        try:
            from openbabel import pybel
            print("检测到 OpenBabel，支持 mol2 格式输出")
        except ImportError:
            print("未安装 OpenBabel（pip install openbabel-wheel 或 conda install -c conda-forge openbabel）")
            print("mol2 格式不可用，已自动回退到 sdf")
            target_format = 'sdf'
            use_openbabel = False

    print(f"正在生成/优化3D结构并转换为 {target_format.upper()} 格式...")

    for i, mol in enumerate(mols):
        m = generate_optimize_3d(mol, names[i])
        m.SetProp("_Name", str(names[i]))

        safe_name = safe_filename(names[i])
        filename = os.path.join(output_dir, f"{safe_name}.{target_format}")

        try:
            if target_format == 'sdf':
                block = Chem.MolToMolBlock(m)
                with open(filename, 'w', encoding='utf-8') as f:
                    f.write(block)
                    f.write('$$$$\n')

            elif target_format == 'mol':
                Chem.MolToMolFile(m, filename)

            elif target_format == 'pdb':
                Chem.MolToPDBFile(m, filename)

            elif target_format == 'mol2':
                # 使用 OpenBabel 的 pybel 进行可靠的 mol2 写入
                sdf_block = Chem.MolToMolBlock(m) + '$$$$\n'
                py_mol = pybel.readstring("sdf", sdf_block)
                py_mol.title = str(names[i])
                py_mol.write("mol2", filename, overwrite=True)

            print(f"  成功保存: {safe_name}.{target_format}")

        except Exception as e:
            print(f"  写入失败 {filename}: {e}")

# 相似性计算（同前）
def compute_similarity(mols, names, output_dir):
    if len(mols) <= 1:
        print("分子数量 ≤1，跳过相似性计算")
        return

    print("正在计算分子间相似性（Morgan指纹，Tanimoto）...")
    fps = [AllChem.GetMorganFingerprintAsBitVect(m, 2, nBits=2048) for m in mols]
    n = len(mols)
    sim_matrix = pd.DataFrame(1.0, index=names, columns=names)

    for i in range(n):
        for j in range(i + 1, n):
            sim = DataStructs.TanimotoSimilarity(fps[i], fps[j])
            sim_matrix.iloc[i, j] = sim_matrix.iloc[j, i] = round(sim, 4)

    out_path = os.path.join(output_dir, "相似性矩阵.xlsx")
    sim_matrix.to_excel(out_path)
    print(f"相似性矩阵保存至：{out_path}")

# main（输入提示更新，支持 mol2）
def main():
    try:
        input_path = input("请输入文件或文件夹路径（支持 xlsx/sdf/mol/mol2/pdb）：").strip().strip('"')
        if not os.path.exists(input_path):
            print("错误：路径不存在")
            return

        all_mols = []
        all_names = []
        all_sources = []

        if os.path.isdir(input_path):
            print("检测到文件夹路径，进入批量模式...")
            base_dir = os.path.abspath(input_path)
            files = [os.path.join(input_path, f) for f in os.listdir(input_path)
                     if os.path.splitext(f)[1].lower() in SUPPORTED_EXT]
            if not files:
                print("文件夹中未找到支持的文件")
                return
            print(f"找到 {len(files)} 个支持的文件")

            for file_path in files:
                mols, mol_names, source = parse_file(file_path)
                all_mols.extend(mols)
                all_names.extend(mol_names)
                all_sources.extend([source] * len(mols))

        else:
            print("单个文件模式")
            base_dir = os.path.dirname(os.path.abspath(input_path))
            mols, mol_names, source = parse_file(input_path)
            all_mols = mols
            all_names = mol_names
            all_sources = [source] * len(mols)

        if not all_mols:
            print("未加载到任何有效分子")
            return

        print(f"总共成功加载 {len(all_mols)} 个分子")

        # 智能名称去重
        name_counts = Counter(all_names)
        if max(name_counts.values(), default=0) > 1:
            print("检测到分子名称重复，自动添加文件名前缀以确保唯一")
            all_names = [f"{os.path.splitext(source)[0]}_{name}" 
                         for name, source in zip(all_names, all_sources)]

        add_source_column = len(set(all_sources)) > 1

        folder_paths = {}
        for eng, chn in FOLDER_MAP.items():
            path = os.path.join(base_dir, eng)
            os.makedirs(path, exist_ok=True)
            folder_paths[chn] = path

        df_props = compute_properties(all_mols, all_names, all_sources, add_source_column)
        props_path = os.path.join(folder_paths["分子信息"], "分子属性.xlsx")
        df_props.to_excel(props_path, index=False)
        print(f"分子属性保存至：{props_path}")

        target = input("请输入目标3D格式（sdf/mol/pdb/mol2，默认sdf）：").strip().lower() or "sdf"
        convert_format(all_mols, all_names, target, folder_paths["格式转换"])
        print(f"格式转换完成，结果在文件夹：{FOLDER_MAP['Format_Conversion']}")

        compute_similarity(all_mols, all_names, folder_paths["相似性比较"])

        print("\n所有任务完成！")

    except Exception as e:
        print(f"\n程序出错：{e}")
        sys.exit(1)


if __name__ == "__main__":
    print("=== 分子信息处理工具（支持 mol2 输出，依赖 OpenBabel） ===")
    print("   author: shenjianlin")
    print("   site: git@github.com:sjl-openmywork/Network-toxicology-script.git\n")
    main()