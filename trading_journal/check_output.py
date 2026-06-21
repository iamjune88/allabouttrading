import sys
sys.stdout.reconfigure(encoding='utf-8')
import pandas as pd

df_fills = pd.read_excel('output/trading_journal.xlsx', sheet_name='Fills')
df_trades = pd.read_excel('output/trading_journal.xlsx', sheet_name='Trades')

print('=== Fills 요약 ===')
print(f'총 {len(df_fills)}건')
print(df_fills.groupby(['거래일자','증권사','종목','방향'])['수량'].sum().to_string())

print('\n=== Trades (라운드트립) ===')
cols = ['거래일자','종목','진입방향','진입수량','진입가격','청산가격','순손익','결과']
print(df_trades[cols].to_string())
print(f'\n순손익 합계: {df_trades["순손익"].sum():,.0f}원')
