import pandas as pd
import urllib

#Grabs list of spreadsheets from Google Docs that are taken from
#treasury financial website. Puts all data into singularly mapped
#spreadsheet with extra information for simple use within Power BI.
class SpreadSheetCombine:
    def __init__(self):
        self.column_data = {};
        self.sheet_column = []
        self.file_column = [];
        self.currency_column = [];
        self.impact_column = [];
        self.df = pd.DataFrame()
    
    def add_file(self, link):
        try:
            urllib.request.urlretrieve(link, "file.xlsx")
            excel_sheet = pd.ExcelFile("file.xlsx")
            
            excel_readable = pd.read_excel("file.xlsx")
            useful_col = excel_readable.columns[1]
            name_list = list(excel_readable[useful_col][1:3]);
            file_name = name_list[0] + "_" + name_list[1];
        except:
            print("Failed to use file: " + link);
            print("Continuing...");
            return;
            
        for sheet_name in excel_sheet.sheet_names:
            self.add_sheet(excel_sheet, file_name, sheet_name)
            
            
    def add_sheet(self, excel_sheet, file_name, sheet_name):
        sheet_df = excel_sheet.parse(sheet_name);
        
        if "Info" in sheet_name:
            return;
        
        for i in range(0,len(sheet_df.index)):
            self.sheet_column.append(sheet_name.replace(" ", ""));
            self.file_column.append(file_name)
            self.impact_column.append(file_name[4:].replace(".xlsx", ""));
            self.currency_column.append(file_name[0:3]);
            
        for col in sheet_df.columns:
            if col in self.column_data.keys():
                self.column_data[col].extend(list(sheet_df[col]))
            else:
                self.column_data[col] = list(sheet_df[col])
        
    
    def create_df_file(self):
        for col in self.column_data.keys():
            self.df[col] = pd.Series(self.column_data[col])
        self.df["SheetName"] = pd.Series(self.sheet_column);
        self.df["FileName"] = pd.Series(self.file_column)
        self.df["Impact"] = pd.Series(self.impact_column)
        self.df["Currency"] = pd.Series(self.currency_column)
        
        writer = pd.ExcelWriter("results.xlsx", engine='xlsxwriter');
        self.df.to_excel(writer);
        writer.save();
    
ssc = SpreadSheetCombine()

#Links redacted - originally from Google Sheets
links=[]

for link in links:
    ssc.add_file(link)

ssc.create_df_file()
