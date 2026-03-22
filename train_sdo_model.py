#!/usr/bin/env python3
"""
ALY 6980 Capstone - Advanced Machine Learning Pipeline
Massachusetts Open Checkbook: Predictive Modeling of SDO Commitment

This script demonstrates a full end-to-end Machine Learning pipeline using scikit-learn:
1. Data Preprocessing (Imputation, One-Hot Encoding)
2. Model Training (Random Forest Regressor)
3. Model Evaluation (RMSE, R2)
4. Feature Importance Visualization and Model Persistence

Author: Sumesh Chakkaravarthi
"""

import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
import seaborn as sns
from pathlib import Path

from sklearn.model_selection import train_test_split
from sklearn.pipeline import Pipeline
from sklearn.compose import ColumnTransformer
from sklearn.preprocessing import OneHotEncoder, StandardScaler
from sklearn.ensemble import RandomForestRegressor
from sklearn.metrics import mean_squared_error, r2_score
import joblib

import warnings
warnings.filterwarnings('ignore')

# ──────────────────────────────────────────────────────────────────
# Configuration
# ──────────────────────────────────────────────────────────────────
VENDOR_FILE = "/Users/sumesh/Downloads/Copy of Vendor Contact Details (1).xlsx"
OUTPUT_DIR = Path("/Users/sumesh/Projects/Antigravity/Capstone/output")
MODEL_DIR = Path("/Users/sumesh/Projects/Antigravity/Capstone/models")

OUTPUT_DIR.mkdir(exist_ok=True, parents=True)
MODEL_DIR.mkdir(exist_ok=True, parents=True)

SKIP_SHEETS = ['Abbreviations ']

plt.rcParams.update({
    'figure.facecolor': '#FAFAFA',
    'axes.facecolor': '#FAFAFA',
    'font.family': 'sans-serif',
    'font.size': 11,
})

# ──────────────────────────────────────────────────────────────────
# Data Loading (Similar to EDA)
# ──────────────────────────────────────────────────────────────────
def load_data(filepath: str) -> pd.DataFrame:
    print("Loading raw vendor data for ML model...")
    xl = pd.ExcelFile(filepath)
    frames = []
    
    for sheet in xl.sheet_names:
        if sheet.strip() in [s.strip() for s in SKIP_SHEETS]:
            continue
            
        df = pd.read_excel(filepath, sheet_name=sheet, header=0)
        df = df.iloc[:, :7]
        df.columns = ['Contract_Code', 'Name', 'Company', 'Role', 'Email', 'Phone', 'SDO_Commitment_Pct']
        df['Category'] = sheet.strip()
        frames.append(df)
        
    combined = pd.concat(frames, ignore_index=True)
    
    # Target Variable: Fill missing SDO with 0% assuming no commitment
    combined['SDO_Commitment_Pct'] = pd.to_numeric(combined['SDO_Commitment_Pct'], errors='coerce').fillna(0)
    
    # Feature Engineering
    combined['Role'] = combined['Role'].astype(str).str.replace('\n', ' ', regex=False).str.strip()
    
    # Clean Role - Group sparse roles into "Other"
    role_counts = combined['Role'].value_counts()
    valid_roles = role_counts[role_counts > 5].index
    combined['Role_Cleaned'] = combined['Role'].apply(lambda x: x if x in valid_roles else 'Other')
    
    print(f"Data loaded: {len(combined)} rows. Ready for training.")
    return combined

# ──────────────────────────────────────────────────────────────────
# Machine Learning Pipeline
# ──────────────────────────────────────────────────────────────────
def train_and_evaluate(df: pd.DataFrame):
    print("\n--- Initializing Machine Learning Pipeline ---")
    
    # Features and Target
    # We predict the SDO % based on Category and the Contact's Role inside the vendor
    features = ['Category', 'Role_Cleaned']
    target = 'SDO_Commitment_Pct'
    
    X = df[features]
    y = df[target] * 100 # Predict percentage as whole number for readability
    
    # Train-test split
    X_train, X_test, y_train, y_test = train_test_split(X, y, test_size=0.2, random_state=42)
    print(f"Training shapes: X={X_train.shape}, y={y_train.shape}")
    print(f"Testing shapes : X={X_test.shape}, y={y_test.shape}")
    
    # Build Scikit-Learn Pipeline
    categorical_features = ['Category', 'Role_Cleaned']
    categorical_transformer = OneHotEncoder(handle_unknown='ignore')
    
    preprocessor = ColumnTransformer(
        transformers=[
            ('cat', categorical_transformer, categorical_features)
        ])
        
    model = Pipeline(steps=[
        ('preprocessor', preprocessor),
        ('regressor', RandomForestRegressor(n_estimators=100, max_depth=10, random_state=42, n_jobs=-1))
    ])
    
    # Train
    print("Training Random Forest Regressor...")
    model.fit(X_train, y_train)
    
    # Evaluate
    y_pred = model.predict(X_test)
    mse = mean_squared_error(y_test, y_pred)
    rmse = np.sqrt(mse)
    r2 = r2_score(y_test, y_pred)
    
    print("\n--- Model Evaluation ---")
    print(f"Root Mean Squared Error (RMSE): {rmse:.2f}%")
    print(f"R-squared (R2) Score        : {r2:.4f}")
    
    # Extract Feature Importances
    preprocessor = model.named_steps['preprocessor']
    regressor = model.named_steps['regressor']
    
    # Get feature names from OneHotEncoder
    ohe_features = preprocessor.named_transformers_['cat'].get_feature_names_out(categorical_features)
    importances = regressor.feature_importances_
    
    feature_imp_df = pd.DataFrame({
        'Feature': [f.split('_')[-1] for f in ohe_features],
        'Importance': importances
    }).sort_values(by='Importance', ascending=False).head(15)
    
    # Plot Feature Importance
    plot_feature_importance(feature_imp_df)
    
    # Save Model
    model_path = MODEL_DIR / 'sdo_rf_model.pkl'
    joblib.dump(model, model_path)
    print(f"Model saved to {model_path}")

def plot_feature_importance(df: pd.DataFrame):
    fig, ax = plt.subplots(figsize=(10, 6))
    
    sns.barplot(x='Importance', y='Feature', data=df, palette='Blues_r', ax=ax)
    
    ax.set_title('Machine Learning: Top 15 Features Predicting SDO Commitment', pad=15, fontweight='bold')
    ax.set_xlabel('Relative Feature Importance', fontweight='bold')
    ax.set_ylabel('Vendor Characteristic (Category / Role)', fontweight='bold')
    
    for i, v in enumerate(df['Importance']):
        ax.text(v + 0.005, i, f"{v:.3f}", va='center', fontsize=9, fontweight='bold', color='#333')
        
    fig.tight_layout()
    output_path = OUTPUT_DIR / 'ML_Feature_Importance.png'
    fig.savefig(output_path, dpi=200)
    print(f"Feature importance plot saved to {output_path}")

# ──────────────────────────────────────────────────────────────────
# Execution
# ──────────────────────────────────────────────────────────────────
if __name__ == "__main__":
    vendor_data = load_data(VENDOR_FILE)
    train_and_evaluate(vendor_data)
    print("Pipeline Execution Completed Successfully.")
