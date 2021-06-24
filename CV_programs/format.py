def txt(source):
    import pandas as pd
    from CV_programs import txt
    read_text_file = pd.read_csv(source)
    read_text_file.to_excel (r"test.xlsx", index=None)
    txt.txt()
def csv(source):
    import pandas as pd
    read_text_file = pd.read_csv(source,delimiter=',')
    read_text_file.to_excel (r"test.xlsx", index=None)
def vba(source):
    import pandas as pd
    read_text_file = pd.read_csv(source,delimiter='\t')
    read_text_file.to_excel (r"test.xlsx", index=None)
