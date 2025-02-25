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

##Perform RFM Analysis
# Prepare RFM DataFrame
rfm_df = df_cleaned[['Customer ID', 'Total Spend', 'Items Purchased', 'Days Since Last Purchase']].copy()

# Rename columns for clarity
rfm_df.rename(columns={'Total Spend': 'Monetary', 'Items Purchased': 'Frequency', 'Days Since Last Purchase': 'Recency'}, inplace=True)

# Following added 2/24/2025 by Julian VB
# Assign RFM Scores
# Recency: lower is better, Frequency and Monetary: higher is better
rfm_df['R_Score'] = pd.qcut(rfm_df['Recency'], q=4, labels=[4, 3, 2, 1]).astype(int)
rfm_df['F_Score'] = pd.qcut(rfm_df['Frequency'], q=4, labels=[1, 2, 3, 4]).astype(int)
rfm_df['M_Score'] = pd.qcut(rfm_df['Monetary'], q=4, labels=[1, 2, 3, 4]).astype(int)

# Create combined RFM score
rfm_df['RFM_Score'] = rfm_df['R_Score'].astype(str) + rfm_df['F_Score'].astype(str) + rfm_df['M_Score'].astype(str)

# Define customer segments based on RFM Score
def categorize_customer(rfm_score):
    if rfm_score in ['444', '433', '434', '343', '344']:
        return 'VIP Customer'
    elif rfm_score in ['311', '211', '111', '121', '131']:
        return 'At-Risk'
    elif rfm_score in ['441', '431', '421', '331', '321']:
        return 'Regular Customer'
    else:
        return 'One-Time Buyer'

# Apply categorization
rfm_df['Customer Segment'] = rfm_df['RFM_Score'].apply(categorize_customer)

# Merge RFM results back into cleaned data
final_df = pd.merge(df_cleaned, rfm_df[['Customer ID', 'RFM_Score', 'Customer Segment']], on='Customer ID', how='left')

##Final Data Check
print("Final dataset shape:", final_df.shape)
print("Remaining missing values:\n", final_df.isnull().sum())

# Save the final cleaned and analyzed dataset
final_file_path = r"C:\DataProjects\final_ecommerce_analysis.xlsx"
final_df.to_excel(final_file_path, index=False)

# Print message confirming the file was saved
print(f"Cleaned and analyzed data saved to {final_file_path}")