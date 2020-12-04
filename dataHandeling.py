import sys

import pandas as pd


def analyze_report(full_path):
    try:
        full_file = pd.read_excel(full_path, header=None)
    except ValueError as e:
        # print(e)
        print('Failed to read file', full_path, sep=' ')
        return None
    except:
        return None
    try:
        df = extract_diff_table(full_file)
    except IndexError as e:
        # print(e)
        print('Did not find strings used to parse file', full_path, sep=' ')
        return None
    except:
        return None
    try:
        return find_diffs(df)
    except:
        print(sys.exc_info()[0])
        return None


def extract_diff_table(full_file):
    """Finds actual DNA data in output spreadsheet,
        This function relies on the fact that the first row for this section has 'First base in Codon' and that the
         last has 'Total Differences'
         """
    yi = full_file.index[full_file[0] == 'First base in Codon'][0] # find the first row
    yj = full_file.index[full_file[0] == 'Total Differences'][0] # find last row
    xi = full_file.loc[yi].last_valid_index() # find last column
    df = full_file.loc[yi:yj, :xi]
    new_header = df.iloc[0]  # grab the first row for the header
    df = df[1:]  # take the data less the header row
    df.columns = new_header  # set the header row as the df header
    return df.reset_index(drop=True)


def find_diffs(df):
    """Goes over each amino acid and find differences, returns dict of diffs"""
    amino_index = df.columns.get_loc('Amino Acid')
    ground_truth = df.iloc[:, amino_index:(amino_index + 2)]
    all_acids = list(df.columns)[amino_index+2:-1:2]
    all_difs = {}
    # go over each amino acid and find diffs
    for acid in all_acids:
        this_acid_diff = []
        current_table = df[acid][1:-1]
        # go over each row in amino acid and check if equal
        for index, current_row in current_table.iterrows():
            if ground_truth.iloc[index, 1] != current_row.iloc[1] and current_row.iloc[1] != 'X':
                this_acid_diff.append((str(ground_truth.iloc[index, 1]) + str(ground_truth.iloc[index, 0]) + str(current_row.iloc[1])))
        # do a sanity check regarding the number of expected differences
        expected_diff = df[acid].tail(1).iloc[0,1]
        if len(this_acid_diff) != expected_diff:
            print('Inconsistency found in', acid, 'found', len(this_acid_diff), 'differences, expected', expected_diff, sep=' ')
            this_acid_diff.append('ERROR!!!!!!!!!')
        if len(this_acid_diff) > 0:
            all_difs[acid] = ','.join(this_acid_diff)
    return pd.DataFrame.from_dict(all_difs, orient='index')


def finalize(frames, output_path):
    writer = pd.ExcelWriter(output_path, engine='xlsxwriter')
    for frame in frames:
        frame[1].to_excel(writer, sheet_name=frame[0])
    writer.save()


if __name__ == '__main__':
    analyze_report(r'C:\Users\Edanz\Downloads\report.xlsx')