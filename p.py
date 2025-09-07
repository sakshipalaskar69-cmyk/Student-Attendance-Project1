import pandas as pd
import matplotlib.pyplot as plt
import seaborn as sns
import os
import openpyxl

#File Detail Path
INPUT_FILE = "FinalAttendance.xlsx"
OUTPUT_DIR = "Attendance_output"
FIG_DIR = os.path.join(OUTPUT_DIR, "figures")
os.makedirs(FIG_DIR,exist_ok=True)

#load file
df=pd.read_excel(INPUT_FILE)
print(df.head())

# A Data Cleaning Operations
# Check Duplicates and drop
print("After Removing Duplicates Rows")
df_cleaned = df.drop_duplicates()

# Fill or drop null values
print("Missing values per column:")
df = df.fillna({
    "Name": "NOT Assign",               
    "ID": "NULL",             
    "Class": "Null",         
    })

#Rename or format Column name
#print("Rename or format column name")
df.columns = df.columns.str.strip().str.upper().str.replace(" ", "_")
print(df.columns)

#Data B Analysis
#Average
#1.Average of present Student:
print("Average ", df['PRESENT'].mean())

#Min/Max Value
max_attendance = df['PRESENT'].max()
min_attendance = df['ABSENT'].min()

print(f"Maximum Attendance Value: {max_attendance}")
print(f"Minimum Attendance Value: {min_attendance}")

#Filter Data
high_scores = df[df['PRESENT'] > 3]
print(high_scores)

#Save Modified Data
clean_file = os.path.join(OUTPUT_DIR, "clean_FinalAttendance.xlsx")

if not os.path.exists(clean_file):
    df.to_excel(clean_file, index=False)
    print(f"\nCleaned data saved to {clean_file}")
else:
    print("File already exists. Please close it or delete it before running.")

#Create Charts
#Bar
df['CLASS'].value_counts().plot(kind='bar', color='pink')
plt.title("Bar chart:Class Vs Present")
plt.xlabel("CLASS")
plt.ylabel("Present")
plt.tight_layout()
plt.show()

#Pie chart
plt.title('Pie Chart: Class Distribution')
df['ADDRESS'].value_counts().plot(
    kind='pie',
    autopct='%1.1f%%',
    startangle=90,
    ylabel=''  
)
plt.show()

#Line chart
plt.plot(df['ABSENT'], df['PRESENT'], marker='o')
plt.title('Line Chart:Absent Vs Present')
plt.xlabel('ABSENT')
plt.ylabel('PRESENT')
plt.grid(True)
plt.show()

#Scatter plot
sns.scatterplot(x='CLASS', y='PRESENT', data=df)
plt.title('Scatter Chart:Class Vs Present')
plt.xlabel('Class')
plt.ylabel('Address')
plt.xticks(rotation=90)
plt.tight_layout()
plt.show()









