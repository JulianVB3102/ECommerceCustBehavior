#Coded by Julian VB. 2/19/2025
import pandas as pd
import matplotlib.pyplot as plt
from sklearn.model_selection import train_test_split
from sklearn.linear_model import LogisticRegression
from sklearn.metrics import accuracy_score, classification_report
from sklearn.preprocessing import StandardScaler
import seaborn as sns

# Load the dataset
file_path = r"C:\DataProjects\business_ecom_churn_predict\E-commerce Customer Behavior - Sheet1.csv"
final_file_path = r"C:\DataProjects\business_ecom_churn_predict\final_business_ecommerce_analysis.xlsx"

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
final_df = pd.merge(df_cleaned, rfm_df[['Customer ID', 'RFM_Score', 'Customer Segment', 'Recency', 'Frequency', 'Monetary']], on='Customer ID', how='left')

## Feature Engineering
# Calculate average spend per purchase
final_df["Avg Spend Per Purchase"] = final_df["Monetary"] / final_df["Frequency"]
final_df["Avg Spend Per Purchase"] = final_df["Avg Spend Per Purchase"].fillna(0)

# Create high-value customer flag (top 10% spenders)
threshold = final_df["Monetary"].quantile(0.9)
final_df["High-Value Customer"] = final_df["Monetary"].apply(lambda x: 1 if x >= threshold else 0)

## Churn Prediction (For All Customers)
# Create churn label: 1 if Days Since Last Purchase > 45, else 0
final_df['Churn'] = final_df['Days Since Last Purchase'].apply(lambda x: 1 if x > 45 else 0)

# Define features and target
features = final_df[['Recency', 'Frequency', 'Monetary', 'Avg Spend Per Purchase', 'High-Value Customer']]

# Scale features to improve model performance
scaler = StandardScaler()
features_scaled = scaler.fit_transform(features)

# Train logistic regression model with increased iterations
model = LogisticRegression(max_iter=500)
model.fit(features_scaled, final_df['Churn'])

# Predict churn probabilities for ALL customers
final_df["Churn Probability"] = model.predict_proba(features_scaled)[:, 1]

# Save final data with full churn predictions
with pd.ExcelWriter(final_file_path, engine="openpyxl", mode='w') as writer:
    final_df.to_excel(writer, sheet_name='Final Business Data', index=False)

# Print confirmation
print(f"Business-ready churn predictions saved to {final_file_path}")
