#Coded by Julian VB. 2/19/2025
import pandas as pd

# Load the dataset
file_path = r"C:\DataProjects\E-commerce Customer Behavior - Sheet1.csv"

df = pd.read_csv(file_path)

##Handle Missing Values
# Drop columns where over 50% of the data is missing
df = df.dropna(thresh=len(df) * 0.5, axis=1)

# Fill missing numerical values with median
df.fillna(df.median(numeric_only=True), inplace=True)

# Fill missing categorical values with the mode (most frequent value)
for col in df.select_dtypes(include=["object"]).columns:
    df[col] = df[col].fillna(df[col].mode()[0])

##Handle Duplicate Entries
# Flag duplicate rows
df["is_duplicate"] = df.duplicated(keep=False)

# Define conditions for removing bad duplicates (e.g., unrealistic values)
invalid_dupes = df[(df["is_duplicate"]) & 
                   ((df["Total Spend"] < 0) |  # Invalid spend
                    (df["Age"] < 18) |  # Unrealistic age
                    (df["Days Since Last Purchase"] < 0))]  # Impossible purchase history

# Remove only erroneous duplicates
df_cleaned = df.drop(invalid_dupes.index)

##Normalize Data
# Convert text fields to lowercase and strip whitespace
text_cols = ["Gender", "City", "Membership Type", "Satisfaction Level"]
df_cleaned[text_cols] = df_cleaned[text_cols].apply(lambda x: x.str.lower().str.strip())

# Convert categorical values to proper formats
df_cleaned["Discount Applied"] = df_cleaned["Discount Applied"].astype(bool)

# Drop the duplicate flag column (not needed in final data)
df_cleaned.drop(columns=["is_duplicate"], inplace=True)

##Final Data Check
print("Final dataset shape:", df_cleaned.shape)
print("Remaining missing values:\n", df_cleaned.isnull().sum())

# Save the cleaned dataset
cleaned_file_path = r"C:\DataProjects\cleaned_ecommerce_data.xlsx"
df_cleaned.to_excel(cleaned_file_path, index=False)

# Print message confirming the file was saved
print(f"Cleaned data saved to {cleaned_file_path}")
