def main():
    # 기존 코드

    import pandas as pd
    from openpyxl import load_workbook

    # Set the working directory and file paths
    working_directory = r'C:/Users/windo/OneDrive/바탕 화면/Zenbook2/DOWIN/inventory/'
    memo_file_path = f'{working_directory}재고수불부-메모사항.xlsx'
    pivot_file_path = f'{working_directory}재고수불부-sheet_pivot0.xlsx'

    # Load the memo data
    memo_df = pd.read_excel(memo_file_path, sheet_name='메모사항')

    # Load the pivot data
    pivot_df = pd.read_excel(pivot_file_path, sheet_name='원판재고')

    # Adjust the length of memo_df to match pivot_df
    if len(memo_df) < len(pivot_df):
        # Add empty strings to the memo_df
        additional_rows = len(pivot_df) - len(memo_df)
        empty_rows = pd.DataFrame([''] * additional_rows, columns=['메모'])
        memo_df = pd.concat([memo_df, empty_rows], ignore_index=True)

    # Add the memo column to the pivot data
    pivot_df['메모'] = memo_df['메모']

    # Save the updated pivot data back to the Excel file
    output_file_path = f'{working_directory}재고수불부-sheet_pivot.xlsx'
    pivot_df.to_excel(output_file_path, sheet_name='원판재고', index=False)

    print(f"파일이 저장되었습니다: {output_file_path}")


if __name__ == "__main__":
    main()
