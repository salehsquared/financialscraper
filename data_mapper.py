from datetime import date, timedelta
import pandas as pd
import sys
    
#Load in a number of excel files from online by iterating through files on the dept. of treasury
#website, starting at a certain date and going until the present day. Takes all separate excel
#files and concatenates data into succinct data mapping.
class ExcelLoader(object):
    def __init__(self):
        self.series_num = 1; #Our first series number (i.e. Series#1)
        self.series_number_dictionary = {}; #Maps Series#X to numerical data.
        self.series_dictionary = {}; #Maps Val:Val:Val to Series#X
        self.dates = [];
        self.months = [];
        self.years = [];
        
    
    #Append data to our dictionaries
    def append_data(self, reader): #reader is our excel sheet
        df = reader
        columns = df.columns;
        
        dfs = dict(tuple(df.groupby(columns[0]))) #group by "I", "II", etc...
        
        for key in dfs.keys():
            data = dfs[key];
            intro = "";
            ndx = 0;
            for index,row in data.iterrows():
                if(ndx < 2 or ":" in row[1]): #Create the subcategory that our keys are in
                    intro = ":" + row[1].replace(" ", "").replace(":","") + intro;
                else:
                    series_key = row[1].replace(" ", "") + intro; #Generate our key
                    if series_key in self.series_dictionary: #Check if our key is already present
                        series_value = self.series_dictionary[series_key]; #Get our Series#X
                        self.series_number_dictionary[series_value].append(row[4]); #Put our value in the list
                    else:
                        series_value = "Series#" + str(self.series_num); #Create our Series#X
                        self.series_dictionary[series_key] = series_value; #Put Series#X in our dictionary
                        self.series_number_dictionary[series_value] = [row[4]]; #Add a new value associated with Series#X
                        self.series_num += 1; #Increase our series # (so we don't have multiple overlapping series)
                ndx += 1;
    
    #Some links aren't available on the treasury website.
    def readsheet(self, link):
        try:
            excel_sheet = pd.read_excel(link);
            self.append_data(excel_sheet)
        except:
            #print("No data found for: " + link);
            #print(sys.exc_info())
            return(False)
        return(True);
        
#Used for testing
#    def print_dict(self):
#        index = 0;
#        for k,v in self.series_dictionary.items():
#            if(index > 10):
#                break;
#            index += 1;
#            print(k, "--->", v);
#
#        index=0;
#        for k,v in self.series_number_dictionary.items():
#            if(index > 10):
#                break;
#            index += 1;
#            print(k, "--->", v);
    
    #Loop through all sites, grabbing data
    def grab_data(self):
        start_date = date(2020, 2, 1) #Start on February 1st, 2020
        end_date = date.today()
        curr_date = start_date;

        #Iterate through all dates, find the associated links, and pull the data
        while curr_date <= end_date:
            curr_month = str(curr_date.month);
            if (len(curr_month) < 2):
                curr_month = "0" + curr_month;
            curr_year = str(curr_date.year)[-2:];
            curr_day = str(curr_date.day);
            if(len(curr_day) < 2):
                curr_day = "0" + curr_day;
            link = "https://fsapps.fiscal.treasury.gov/dts/files/" + curr_year + curr_month + curr_day + "00" + ".xlsx";
            if(self.readsheet(link)):
                self.dates.append(curr_year + "-" + curr_month + "-" + curr_day);
                self.years.append(curr_date.year);
                self.months.append(curr_date.month);
            curr_date += timedelta(days=1)
    
    #Place data from our dictionaries into an excel file with a new mapping
    def reload_to_excel(self):
        writer = pd.ExcelWriter('data_mapping.xlsx', engine='xlsxwriter');

        complete_mapping = [list(self.series_dictionary.values()), list(self.series_dictionary.keys())]
        
        spaces = [""] * 30;
        cols = ["Series#", "Mapping"];
        cols.extend(spaces)
        cols.extend(["Date", "Year", "Month"])
        cols.extend(list(self.series_dictionary.values()));
        
        complete_mapping = pd.DataFrame(columns=cols)
        
        #Mapping for Series# ---> X:Y:Z
        complete_mapping["Series#"] = list(self.series_dictionary.values());
        complete_mapping["Mapping"] = list(self.series_dictionary.keys());
        
        complete_mapping["Date"] = pd.Series(list(self.dates));
        complete_mapping["Year"] = pd.Series(list(self.years));
        complete_mapping["Month"] = pd.Series(list(self.months));
        
        series_values = list(self.series_dictionary.values())
        for i in range(len(series_values)):
            complete_mapping[series_values[i]] = pd.Series(list(self.series_number_dictionary[series_values[i]]))
            
        complete_mapping.to_excel(writer);
        writer.save()

el = ExcelLoader();
el.grab_data();
#el.print_dict();
el.reload_to_excel();
        
    


