def main():
    # 기존 코드

    import pandas as pd

    # Load the Excel file
    file_path = r'C:/Users/windo/OneDrive/바탕 화면/Zenbook2/DOWIN/inventory/재고수불부-rmpw.xlsx'
    sheet_name = '원판재고'

    # Read the data from the specified sheet, starting from the 6th row to include data from 7th row
    df = pd.read_excel(file_path, sheet_name=sheet_name, skiprows=5)

    # Select columns A to AD (0-indexed, so 30 columns)
    df = df.iloc[:, :30]

    # Fill merged cell data by forward filling, excluding columns M, N, O, P
    cols_to_fill = df.columns[:12].tolist() + df.columns[16:17].tolist()  # Excluding M, N, O, P

    # Forward fill for specified columns
    df[cols_to_fill] = df[cols_to_fill].ffill()

    # Drop rows where all columns have NaN values
    df.dropna(how='all', inplace=True)

    # Remove rows containing '합          계'
    df = df[~df.apply(lambda row: row.astype(str).str.contains('합          계').any(), axis=1)]

    # Rename columns as specified
    new_columns = ['메이커', '품명', '색상및구조', '두께(㎜)', '폭실측(㎜)', '폭(㎜)', '길이(㎜)', '전월재고(장)', '전월재고(㎡)', 
                '현재고(장)', '현재고(㎡)', '가용재고(㎡)', '입고예정일', '입고(장)', '입고(㎡)', '입항예정일', 'REMARK', 
                '서울당월현장', '서울당월수량(㎡)', '서울+1월현장', '서울+1월수량(㎡)', '서울+2월현장', '서울+2월수량(㎡)', 
                '대구당월현장', '대구당월수량(㎡)', '대구+1월현장', '대구+1월수량(㎡)', '대구+2월현장', '대구+2월수량(㎡)', '메모']
    df.columns = new_columns

    # Define a function to apply the specified transformations
    def clean_data(df):
        # Renaming specific values as per the user's instructions
        replacements = {
            '폴  리  갈': '폴리갈',
            '판   넬   웰': '판넬웰',
            'U 판 넬 (폴 리 유)': '폴리유',
            '기   타': '기타',
            '굿라이프(GL)': '굿라이프',
            '폴리(POLYEE)': '폴리',
            '使用不可 : 입고시기 미상. 판재 상태 심히 불량함. 폐기 대상임!': '폐기대상'
        }
        
        df.replace(replacements, inplace=True)
        
        return df

    # Apply the cleaning function to the dataframe
    df = clean_data(df)

    # Remove 'alt-enter' (newline) characters from '메모' column
    df['메모'] = df['메모'].fillna('').astype(str).str.replace('\n', ' ', regex=False)

    # Change date format in '입고예정일' (column M) and '입항예정일' (column P)
    def change_date_format(date_series):
        return pd.to_datetime(date_series, errors='coerce').dt.strftime('%y%m%d')

    # Preserve the original data for M, N, O, P columns before date format conversion
    original_columns = df[['입고예정일', '입고(장)', '입고(㎡)', '입항예정일']]

    df['입고예정일'] = change_date_format(df['입고예정일'])
    df['입항예정일'] = change_date_format(df['입항예정일'])

    # Convert the columns to text format to retain leading zeros
    df['입고예정일'] = df['입고예정일'].astype(str)
    df['입항예정일'] = df['입항예정일'].astype(str)

    # Replace "nan" with empty string in '입고예정일' and '입항예정일'
    df['입고예정일'].replace('nan', '', inplace=True)
    df['입항예정일'].replace('nan', '', inplace=True)

    # Restore the original data for N, O columns
    df[['입고(장)', '입고(㎡)']] = original_columns[['입고(장)', '입고(㎡)']]

    # Save the cleaned dataframe to a new Excel file in the specified directory and filename
    output_file_path = r'C:/Users/windo/OneDrive/바탕 화면/Zenbook2/DOWIN/inventory/재고수불부-sheet_pivot0.xlsx'
    df.to_excel(output_file_path, sheet_name='원판재고', index=False)

    print(f"파일이 저장되었습니다: {output_file_path}")


if __name__ == "__main__":
    main()
