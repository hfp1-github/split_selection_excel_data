import pandas as pd
import xlwings as xw
import os

def extract_and_save_data_label(df_source, savedir, prefix, suffix=''):
    df_data = df_source.iloc[:,0:-1]
    df_label = df_source.iloc[:,[0,-1]]

    # save data
    df_data.to_csv(os.path.join(savedir, '{}_data{}.csv'.format(prefix, suffix)), header=False, index=False)
    df_label.to_csv(os.path.join(savedir, '{}_label{}.csv'.format(prefix, suffix)), header=False, index=False)

def split_selection_excel_data(tarin_last_row_num, vali_last_row_num, test_last_row_num=None, suffix=''):
    workbooks = xw.books
    wb = workbooks.active

    wb_dir = os.path.dirname(wb.fullname)
    wb_name = os.path.splitext(wb.name)[0]

    cells = wb.selection
    timestamp_cells = cells[:,0]
    timestamp_cells.autofit()
    timestamps_text = [s.api.text for s in timestamp_cells]

    df = pd.DataFrame(cells.value)
    df.iloc[:,0] = timestamps_text # label column cast to int
    df.iloc[:,-1] = df.iloc[:,-1].astype(int) # label column cast to int

    if not test_last_row_num:
        test_last_row_num = len(df)

    df_train = df.iloc[:tarin_last_row_num+1]
    df_vali = df.iloc[tarin_last_row_num+1: vali_last_row_num+1]
    df_test = df.iloc[vali_last_row_num+1: test_last_row_num+1]

    extract_and_save_data_label(df_train, wb_dir, wb_name + '_train', '_1')
    extract_and_save_data_label(df_vali, wb_dir, wb_name + '_vali', '_1')
    extract_and_save_data_label(df_test, wb_dir, wb_name + '_test', '_1')

import argparse

parser = argparse.ArgumentParser(
    formatter_class=argparse.RawDescriptionHelpFormatter,
    description=\
    '''
    excelのデータからtrain, vali, testのデータを作成。
    データソースとしたいセルをexcel上で選択した状態で実行。
    '''
    )

parser.add_argument('train_last_row_num', help='訓練データの最終行番号') 
parser.add_argument('vali_last_row_num', help='バリデーションデータの最終行番号')
parser.add_argument('-t', '--test_last_row_num', help='テストデータの最終行番号', default=None)
parser.add_argument('-s', '--suffix', help='出力ファイル名のsuffix', default='') 

args = parser.parse_args()

if args.test_last_row_num:
    args.test_last_row_num = int(args.test_last_row_num)

split_selection_excel_data(
    int(args.train_last_row_num),
    int(args.vali_last_row_num),
    args.test_last_row_num,
    args.suffix
    )