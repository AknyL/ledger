#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
记账助手 · 自动导入脚本（auto_import.py）

作用：把 bills/ 文件夹里的微信/支付宝账单文件（.csv / .xlsx）解析、去重、
      合并成 ledger.json。应用打开时会自动「同步」这个文件，实现“放进文件夹
      → 自动记账”的闭环。

用法（在 记账助手 文件夹里，用 Python 运行）：
    python auto_import.py            增量处理 bills/ 里新增/变动的文件
    python auto_import.py --force    清空已处理记录，全量重建 ledger.json
    python auto_import.py --watch    持续监听 bills/（每 5 秒扫一次，Ctrl+C 停止）

说明：
  - 本脚本只解析“原始流水”，分类规则与去重由应用在同步时统一处理，
    所以应用里的自定义分类、预算等设置对导入的数据同样生效。
  - 已经处理过的文件会记录在 bills/.processed.json，重复运行不会重复计入。
  - 删除某个账单文件后想重算，用 --force 全量重建。
"""
import os, sys, json, csv, re, io, time, datetime

# Windows 控制台默认可能是 GBK，强制 UTF-8 输出避免中文乱码
try:
    sys.stdout.reconfigure(encoding='utf-8')
    sys.stderr.reconfigure(encoding='utf-8')
except Exception:
    pass

HERE = os.path.dirname(os.path.abspath(__file__))
BILLS = os.path.join(HERE, 'bills')
LEDGER = os.path.join(HERE, 'ledger.json')
MANIFEST = os.path.join(BILLS, '.processed.json')

WATCH_INTERVAL = 5  # 秒


# ---------- 工具 ----------
def normalize(s):
    return re.sub(r'\s+', '', '' if s is None else str(s))


def parse_amount(s):
    if s is None:
        return None
    t = str(s)
    for ch in ('¥', '￥', ',', ' ', '\t', '元'):
        t = t.replace(ch, '')
    t = re.sub(r'^[^\d.\-]+', '', t).strip()
    if t.startswith('(') and t.endswith(')'):
        t = '-' + t[1:-1]
    try:
        return float(t)
    except ValueError:
        return None


def parse_time(s):
    if s is None:
        return None
    t = str(s).strip()
    if not t:
        return None
    m = re.search(r'(\d{4})[-/](\d{1,2})[-/](\d{1,2})[ T]?(\d{1,2})?:?(\d{1,2})?:?(\d{1,2})?', t)
    if not m:
        return None
    y, mo, d = int(m[1]), int(m[2]), int(m[3])
    h, mi, s = int(m[4] or 0), int(m[5] or 0), int(m[6] or 0)
    try:
        return datetime.datetime(y, mo, d, h, mi, s)
    except ValueError:
        return None


def map_direction(dir_str, type_):
    d = normalize(dir_str)
    if '支出' in d:
        return 'expense'
    if '收入' in d:
        return 'income'
    if '中性' in d or '不计' in d or '其他' in d:
        return 'neutral'
    t = normalize(type_)
    if re.search(r'提现|充值|理财|余额宝|零钱通|还款|赎回|买入|卖出|转入|转出|信用卡', t):
        return 'neutral'
    if re.search(r'退款|收益|利息|收款', t):
        return 'income'
    if re.search(r'消费|支付|缴费', t):
        return 'expense'
    return 'expense'


def tx_sig(tx):
    return '|'.join(str(tx[k]) for k in ('time_ms', 'amount', 'direction', 'merchant', 'goods', 'source'))


def make_raw(dt, amount, direction, type_, merchant, goods, pay, status, note, source):
    """生成一条与 index.html 中 makeTx 兼容的原始流水（分类与 sig 由应用补齐）。"""
    return {
        'time': dt.strftime('%Y-%m-%dT%H:%M:%S'),
        'time_ms': int(dt.timestamp() * 1000),  # 仅脚本内部去重用，不写入 ledger.json
        'amount': amount,
        'direction': direction,
        'type': type_,
        'merchant': merchant or '(未注明)',
        'goods': goods or '',
        'pay': pay or '',
        'status': status or '',
        'note': note or '',
        'source': source,
    }


# ---------- 文件读取 ----------
def read_text(path):
    raw = open(path, 'rb').read()
    for enc in ('utf-8-sig', 'gb18030'):
        try:
            return raw.decode(enc)
        except (UnicodeDecodeError, LookupError):
            continue
    return raw.decode('utf-8', errors='replace')


def read_csv_rows(path):
    text = read_text(path)
    return list(csv.reader(io.StringIO(text)))


def read_xlsx_rows(path):
    try:
        from openpyxl import load_workbook
    except ImportError:
        print('[错误] 未安装 openpyxl，无法读取 .xlsx。可改用 CSV 导出，或安装：pip install openpyxl')
        return []
    wb = load_workbook(path, read_only=True, data_only=True)
    ws = wb[wb.sheetnames[0]]
    rows = []
    for row in ws.iter_rows(values_only=True):
        rows.append(['' if c is None else str(c) for c in row])
    wb.close()
    return rows


# ---------- 解析：微信 ----------
def parse_wechat(rows):
    header_idx = -1
    for i, r in enumerate(rows):
        joined = ','.join(normalize(c) for c in r)
        if ('交易时间' in joined and '交易类型' in joined) or \
           ('交易时间' in joined and '收/支' in joined):
            header_idx = i
            break
    if header_idx < 0:
        return None, '未找到微信账单表头'

    head = [normalize(c) for c in rows[header_idx]]

    def find(*keys):
        for k in keys:
            if k in head:
                return head.index(k)
        return -1

    col = {
        'time': find('交易时间'),
        'type': find('交易类型'),
        'merchant': find('交易对方'),
        'goods': find('商品'),
        'dir': find('收/支', '收支'),
        'amount': find('金额(元)', '金额（元）', '金额'),
        'pay': find('支付方式'),
        'status': find('当前状态', '状态'),
        'note': find('备注'),
    }
    if col['time'] < 0 or col['amount'] < 0:
        return None, '微信账单缺少关键列（交易时间/金额）'

    out = []
    for r in rows[header_idx + 1:]:
        if not r or len(r) < 2:
            continue

        def g(idx):
            return str(r[idx]).strip() if 0 <= idx < len(r) else ''

        time_str, amount_str = g(col['time']), g(col['amount'])
        if not time_str or not amount_str:
            continue
        dt = parse_time(time_str)
        amount = parse_amount(amount_str)
        if not dt or amount is None or amount == 0:
            continue
        dir_str = g(col['dir'])
        type_ = g(col['type'])
        direction = map_direction(dir_str, type_)
        out.append(make_raw(dt, amount, direction, type_, g(col['merchant']), g(col['goods']),
                            g(col['pay']), g(col['status']), g(col['note']), 'wechat'))
    return out, None


# ---------- 解析：支付宝 ----------
def parse_alipay(rows):
    header_idx = -1
    for i, r in enumerate(rows):
        joined = ','.join(normalize(c) for c in r)
        if ('交易号' in joined and '交易对方' in joined) or \
           ('交易号' in joined and '金额' in joined):
            header_idx = i
            break
    if header_idx < 0:
        return None, '未找到支付宝账单表头'

    head = [normalize(c) for c in rows[header_idx]]

    def find(*keys):
        for k in keys:
            if k in head:
                return head.index(k)
        return -1

    col = {
        'time': find('交易创建时间', '付款时间', '交易时间'),
        'type': find('类型'),
        'merchant': find('交易对方'),
        'goods': find('商品名称', '商品'),
        'dir': find('收/支', '收支'),
        'amount': find('金额（元）', '金额(元)', '金额'),
        'pay': find('交易来源地'),
        'status': find('交易状态'),
        'note': find('备注'),
    }
    if col['time'] < 0 or col['amount'] < 0:
        return None, '支付宝账单缺少关键列（时间/金额）'

    out = []
    for r in rows[header_idx + 1:]:
        if not r or len(r) < 2:
            continue

        def g(idx):
            return str(r[idx]).strip() if 0 <= idx < len(r) else ''

        time_str, amount_str = g(col['time']), g(col['amount'])
        if not time_str or not amount_str:
            continue
        dt = parse_time(time_str)
        amount = parse_amount(amount_str)
        if not dt or amount is None or amount == 0:
            continue
        status = g(col['status'])
        if status and re.search(r'关闭|失败|等待|未支付|已撤销|取消', status):
            continue
        type_ = g(col['type'])
        direction = map_direction(g(col['dir']), type_)
        out.append(make_raw(dt, amount, direction, type_, g(col['merchant']), g(col['goods']),
                            g(col['pay']), status, g(col['note']), 'alipay'))
    return out, None


# ---------- 识别并解析 ----------
def detect_and_parse(rows):
    joined = '\n'.join(' '.join(str(c) for c in r) for r in rows[:40])
    any_header = '\n'.join(','.join(normalize(c) for c in r) for r in rows[:40])
    if '微信支付' in joined:
        return parse_wechat(rows)
    if '支付宝' in joined:
        return parse_alipay(rows)
    if '交易类型' in any_header:
        return parse_wechat(rows)
    if '交易号' in any_header:
        return parse_alipay(rows)
    return None, '无法识别账单类型（既不是微信也不是支付宝账单格式）'


def parse_file(path):
    ext = os.path.splitext(path)[1].lower()
    if ext == '.xlsx':
        rows = read_xlsx_rows(path)
    elif ext in ('.csv', '.xls'):
        # .xls 老格式读不了就按 CSV 尝试
        rows = read_csv_rows(path)
    else:
        return [], '不支持的文件类型 ' + ext
    if not rows:
        return [], '文件为空或读取失败'
    return detect_and_parse(rows)


# ---------- 主流程 ----------
def load_json(path):
    if os.path.exists(path):
        try:
            with open(path, 'r', encoding='utf-8') as f:
                return json.load(f)
        except (ValueError, OSError):
            return []
    return []


def save_json(path, data):
    with open(path, 'w', encoding='utf-8') as f:
        json.dump(data, f, ensure_ascii=False, indent=2)


def clean_ledger_item(t):
    """写入 ledger.json 前去掉脚本内部字段 time_ms。"""
    return {k: v for k, v in t.items() if k != 'time_ms'}


def run(force=False):
    os.makedirs(BILLS, exist_ok=True)

    manifest = {} if force else (load_json(MANIFEST) or {})
    ledger = load_json(LEDGER)
    # 已有流水去重用的签名集合（从 ledger 现算）
    existing_sigs = set()

    def sig_of_item(item):
        dt = parse_time(item.get('time'))
        ms = int(dt.timestamp() * 1000) if dt else 0
        return '|'.join(str(item.get(k, '')) for k in ('amount', 'direction', 'merchant', 'goods', 'source')) + '|' + str(ms)

    for item in ledger:
        existing_sigs.add(sig_of_item(item))

    files = [f for f in os.listdir(BILLS) if f.lower().endswith(('.csv', '.xlsx', '.xls'))]
    files.sort()

    added_total = 0
    for fname in files:
        path = os.path.join(BILLS, fname)
        stat = os.stat(path)
        stamp = (int(stat.st_mtime), stat.st_size)
        if not force and manifest.get(fname) == list(stamp):
            continue

        parsed, err = parse_file(path)
        if err:
            print(f'  [跳过] {fname}: {err}')
            continue

        added = 0
        for tx in parsed:
            if sig_of_item(tx) in existing_sigs:
                continue
            ledger.append(clean_ledger_item(tx))
            existing_sigs.add(sig_of_item(tx))
            added += 1
        manifest[fname] = list(stamp)
        added_total += added
        print(f'  [√] {fname}: 新增 {added} 笔')

    # 按时间排序后写回
    def time_key(item):
        dt = parse_time(item.get('time'))
        return dt.timestamp() if dt else 0

    ledger.sort(key=time_key)
    save_json(LEDGER, ledger)
    save_json(MANIFEST, manifest)

    exp = sum(1 for t in ledger if t.get('direction') == 'expense')
    inc = sum(1 for t in ledger if t.get('direction') == 'income')
    print('')
    print(f'  本次新增 {added_total} 笔；ledger.json 共 {len(ledger)} 笔（支出 {exp} / 收入 {inc}）')
    print(f'  打开应用点击「🔄 同步」或刷新页面即可入库。')


def main():
    args = sys.argv[1:]
    force = '--force' in args
    watch = '--watch' in args

    print('==============================================')
    print('  记账助手 · 自动导入')
    print('==============================================')
    print(f'  账单文件夹: {BILLS}')
    print('  把微信/支付宝导出的 .csv 或 .xlsx 放进上面的文件夹，')
    print('  然后运行本脚本（或 --watch 自动监听）。')
    print('')

    if not os.path.exists(BILLS):
        os.makedirs(BILLS, exist_ok=True)

    if watch:
        print(f'  正在监听 bills/ 文件夹（每 {WATCH_INTERVAL} 秒检查一次，Ctrl+C 停止）…')
        print('')
        try:
            while True:
                run(force=False)
                time.sleep(WATCH_INTERVAL)
        except KeyboardInterrupt:
            print('\n  已停止监听。')
    else:
        run(force=force)


if __name__ == '__main__':
    main()
