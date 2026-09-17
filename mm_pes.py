# -*- coding: utf-8 -*-
import sys
import pandas as pd
from pathlib import Path
from datetime import datetime
from MM_VMware import parse_utilization

# Входные файлы
BASE_DIR = Path(__file__).resolve().parent
RSM_FILE = BASE_DIR / 'rsm' / 'c9d0d891_rsm.xlsx'
TARGET_FILE = BASE_DIR / 'data_pes' / '20260917_160818.xlsx'
PU_FILE = BASE_DIR / 'PU' / 'PU.xlsx'
OUTPUT_DIR = BASE_DIR / 'output_pes'
LOG_DIR = BASE_DIR / 'log_pes'

# Граница производственной утилизации (в процентах)
PU_BORDER = 50.0

# Ожидаемые столбцы в каждом из файлов
RSM_REQUIRED_COLUMNS = ['Сервер Имя', 'Сервер IP address', 'Сервер КЭ']
TARGET_REQUIRED_COLUMNS = ['node_name', 'node_state', 'max_cpu', 'cpu_2sigma_usage']
PU_REQUIRED_COLUMNS = ['Сервер КЭ', 'Производственная утилизация']

# Допустимые значения node_state для фильтрации
ALLOWED_NODE_STATES = {'["maintenance"]', '["disconnected"]'}

# Порядок первых столбцов в выходном файле
OUTPUT_COLUMN_ORDER = ['КЭ', 'node_name', 'node_state', 'Производственная утилизация']


def check_file_exists(file_path: str) -> None:
    path = Path(file_path)
    if not path.exists():
        print(f"\n[ОШИБКА] Файл не найден: {file_path}")
        print("ОТКАЗ ОТ ВЫПОЛНЕНИЯ СКРИПТА")
        sys.exit(1)
    print(f"  ✓ Файл найден: {path.name}")


def normalize_columns(df: pd.DataFrame) -> None:
    df.columns = [str(c).strip() for c in df.columns]


def check_required_columns(df: pd.DataFrame, required: list, file_label: str) -> None:
    actual_columns = list(df.columns)
    missing = [col for col in required if col not in actual_columns]
    if missing:
        print(f"\n[ОШИБКА] В файле '{file_label}' отсутствуют обязательные столбцы:")
        for col in missing:
            print(f"  ✗ '{col}'")
        print(f"\n  Фактические столбцы в файле ({len(actual_columns)}):")
        for col in actual_columns:
            print(f"    • '{col}'")
        print("\nПроверьте названия столбцов (возможны опечатки или лишние пробелы).")
        print("ОТКАЗ ОТ ВЫПОЛНЕНИЯ СКРИПТА")
        sys.exit(1)
    print(f"  ✓ Обязательные столбцы найдены: {required}")


def check_file_not_empty(df: pd.DataFrame, file_label: str) -> None:
    if len(df) == 0:
        print(f"\n[ОШИБКА] Файл '{file_label}' пуст (0 строк).")
        print("ОТКАЗ ОТ ВЫПОЛНЕНИЯ СКРИПТА")
        sys.exit(1)
    print(f"  ✓ Файл содержит данные: {len(df)} строк")


def clean_name(name):
    if pd.isna(name) or not name:
        return ''
    name = str(name).strip()
    return name.split('.')[0].lower() if name else ''


def filter_and_deduplicate(df: pd.DataFrame) -> pd.DataFrame:
    initial_count = len(df)
    df_filtered = df[df['node_state'].isin(ALLOWED_NODE_STATES)].copy()
    filtered_by_state_count = len(df_filtered)
    print(f"  После фильтрации по node_state: {filtered_by_state_count} строк "
          f"(отброшено: {initial_count - filtered_by_state_count})")

    if filtered_by_state_count == 0:
        print("\n[ОШИБКА] После фильтрации по node_state не осталось ни одной записи.")
        print("ОТКАЗ ОТ ВЫПОЛНЕНИЯ СКРИПТА")
        sys.exit(1)

    df_filtered['max_cpu'] = pd.to_numeric(df_filtered['max_cpu'], errors='coerce').fillna(0)
    df_filtered['cpu_2sigma_usage'] = pd.to_numeric(df_filtered['cpu_2sigma_usage'], errors='coerce').fillna(0)

    df_result = (
        df_filtered
        .sort_values(['max_cpu', 'cpu_2sigma_usage'], ascending=[False, False])
        .drop_duplicates(subset=['node_name'], keep='first')
    )

    print(f"  После дедупликации по node_name: {len(df_result)} уникальных записей "
          f"(отброшено дубликатов: {filtered_by_state_count - len(df_result)})")
    return df_result


def reorder_columns(df: pd.DataFrame) -> pd.DataFrame:
    all_columns = list(df.columns)
    remaining_columns = [col for col in all_columns if col not in OUTPUT_COLUMN_ORDER]
    new_order = OUTPUT_COLUMN_ORDER + remaining_columns
    print(f"  Итоговый порядок столбцов: {new_order}")
    return df[new_order]


def save_result(df, path):
    path.parent.mkdir(parents=True, exist_ok=True)
    df[OUTPUT_COLUMN_ORDER].to_excel(path, index=False)
    print(f'Сохранение в {path}')


def main(rsm_file=RSM_FILE, target_file=TARGET_FILE, pu_file=PU_FILE,
         border=PU_BORDER, output_dir=OUTPUT_DIR, log_dir=LOG_DIR, run_id=None):
    RSM_FILE, TARGET_FILE, PU_FILE = map(Path, (rsm_file, target_file, pu_file))
    PU_BORDER = float(border)
    if not 0 <= PU_BORDER <= 100:
        raise ValueError('Порог утилизации должен быть от 0 до 100%.')
    if hasattr(sys.stdout, 'reconfigure'):
        sys.stdout.reconfigure(encoding='utf-8')
    # 1. Формируем динамическое имя выходного файла
    target_path = Path(TARGET_FILE)
    timestamp = run_id or datetime.now().strftime("%Y%m%d_%H%M%S_%f")
    output_filename = f"{target_path.stem}_{timestamp}-with-CE.xlsx"
    log_file = Path(log_dir) / output_filename
    output_file = Path(output_dir) / f'MM_{output_filename}'

    # 2. Проверка существования файлов
    print("\n=== ПРОВЕРКА ФАЙЛОВ ===")
    check_file_exists(RSM_FILE)
    check_file_exists(TARGET_FILE)
    check_file_exists(PU_FILE)

    # 3. Чтение и проверка RSM файла
    print("\n=== ЧТЕНИЕ RSM ФАЙЛА ===")
    try:
        df_rsm = pd.read_excel(RSM_FILE)
    except Exception as e:
        print(f"\n[ОШИБКА] Не удалось прочитать файл '{RSM_FILE}': {e}")
        print("ОТКАЗ ОТ ВЫПОЛНЕНИЯ СКРИПТА")
        sys.exit(1)

    normalize_columns(df_rsm)
    check_file_not_empty(df_rsm, Path(RSM_FILE).name)
    check_required_columns(df_rsm, RSM_REQUIRED_COLUMNS, Path(RSM_FILE).name)
    print("\n=== RSM ФАЙЛ поля проверены ===")

    # Создание маппинга CE
    print("\n=== Создаём словарь маппинга: name/ip -> CE ===")
    ce_map = {}
    for idx, row in df_rsm.iterrows():
        server_name = str(row['Сервер Имя']).strip() if pd.notna(row['Сервер Имя']) else ''
        server_ip = str(row['Сервер IP address']).strip() if pd.notna(row['Сервер IP address']) else ''
        ce_value = row['Сервер КЭ']

        if pd.isna(ce_value) or str(ce_value).strip() == '-' or str(ce_value).strip() == '':
            continue

        clean_srv_name = clean_name(server_name)
        if clean_srv_name:
            ce_map[clean_srv_name] = ce_value
        if server_ip and server_ip != '-':
            ce_map[server_ip] = ce_value

        if idx < 5:
            print(f"  name='{server_name}' -> clean='{clean_srv_name}' | ip='{server_ip}' | CE={ce_value}")
    print(f"\n  Маппингов создано: {len(ce_map)}")

    # 4. Чтение и проверка целевого файла
    print("\n=== ЧТЕНИЕ ЦЕЛЕВОГО ФАЙЛА ===")
    try:
        df_target = pd.read_excel(TARGET_FILE)
    except Exception as e:
        print(f"\n[ОШИБКА] Не удалось прочитать файл '{TARGET_FILE}': {e}")
        print("ОТКАЗ ОТ ВЫПОЛНЕНИЯ СКРИПТА")
        sys.exit(1)

    normalize_columns(df_target)
    check_file_not_empty(df_target, target_path.name)
    check_required_columns(df_target, TARGET_REQUIRED_COLUMNS, target_path.name)

    # Добавляем столбец КЭ
    ce_results = []
    matched = 0
    for idx, row in df_target.iterrows():
        matched_ce = None
        node_name = str(row['node_name']) if pd.notna(row['node_name']) else ''
        clean_node_name = clean_name(node_name)

        if clean_node_name and clean_node_name in ce_map:
            matched_ce = ce_map[clean_node_name]
            matched += 1
        elif node_name and node_name in ce_map:
            matched_ce = ce_map[node_name]
            matched += 1
        ce_results.append(matched_ce)

    df_target['КЭ'] = ce_results
    print(f"\nНайдено совпадений по КЭ: {matched} из {len(df_target)} "
          f"({matched / len(df_target) * 100:.1f}%)")

    # 5. Фильтрация и дедупликация
    print("\n=== ФИЛЬТРАЦИЯ И ДЕДУПЛИКАЦИЯ ===")
    df_result = filter_and_deduplicate(df_target)

    # 6. Чтение и слияние с файлом PU
    print("\n=== ОБЪЕДИНЕНИЕ С ФАЙЛОМ PU ===")
    try:
        df_pu = pd.read_excel(PU_FILE)
    except Exception as e:
        print(f"\n[ОШИБКА] Не удалось прочитать файл '{PU_FILE}': {e}")
        print("ОТКАЗ ОТ ВЫПОЛНЕНИЯ СКРИПТА")
        sys.exit(1)

    normalize_columns(df_pu)
    check_file_not_empty(df_pu, Path(PU_FILE).name)
    check_required_columns(df_pu, PU_REQUIRED_COLUMNS, Path(PU_FILE).name)

    # Создаём маппинг: Сервер КЭ -> Производственная утилизация
    pu_map = {}
    for _, row in df_pu.iterrows():
        ce_key = row['Сервер КЭ']
        if pd.notna(ce_key) and str(ce_key).strip():
            pu_map[str(ce_key).strip()] = row['Производственная утилизация']

    # Применяем маппинг
    df_result['Производственная утилизация'] = df_result['КЭ'].map(
        lambda x: pu_map.get(str(x).strip()) if pd.notna(x) else None
    )
    print(f"  ✓ Поле 'Производственная утилизация' добавлено.")

    # Снимок до применения border.
    save_result(df_result, log_file)

    # 7. Применение границы утилизации (border)
    print(f"\n=== ПРИМЕНЕНИЕ ГРАНИЦЫ УТИЛИЗАЦИИ ({PU_BORDER}%) ===")
    df_result['_util_num'] = parse_utilization(df_result['Производственная утилизация'])

    before_count = len(df_result)
    # Явная маска: оставляем строки, где утилизация <= border ИЛИ значение отсутствует (NaN)
    keep_mask = (df_result['_util_num'] <= PU_BORDER) | df_result['_util_num'].isna()
    df_result = df_result[keep_mask].drop(columns=['_util_num']).copy()

    print(f"  Осталось записей: {len(df_result)} (удалено по условию border: {before_count - len(df_result)})")

    # 8. Переупорядочивание столбцов
    print("\n=== ПЕРЕУПОРЯДОЧИВАНИЕ СТОЛБЦОВ ===")
    df_result = reorder_columns(df_result)

    # 9. Сохранение
    print(f"\n=== СОХРАНЕНИЕ РЕЗУЛЬТАТА ===")
    print(f"Сохранение в {output_file}...")
    try:
        save_result(df_result, output_file)
    except Exception as e:
        print(f"\n[ОШИБКА] Не удалось сохранить файл '{output_file}': {e}")
        print("Возможно, файл открыт в Excel или нет прав на запись.")
        print("ОТКАЗ ОТ ВЫПОЛНЕНИЯ СКРИПТА")
        sys.exit(1)
    print("✓ Готово!")


if __name__ == '__main__':
    import argparse

    parser = argparse.ArgumentParser(description='Обработка PES')
    parser.add_argument('--rsm', default=RSM_FILE)
    parser.add_argument('--target', default=TARGET_FILE)
    parser.add_argument('--pu', default=PU_FILE)
    parser.add_argument('--border', type=float, default=PU_BORDER)
    parser.add_argument('--output-dir', default=OUTPUT_DIR)
    parser.add_argument('--log-dir', default=LOG_DIR)
    parser.add_argument('--run-id')
    args = parser.parse_args()
    main(args.rsm, args.target, args.pu, args.border, args.output_dir, args.log_dir, args.run_id)
