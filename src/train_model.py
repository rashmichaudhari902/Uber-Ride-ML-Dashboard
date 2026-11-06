"""
train_model.py - Machine Learning Pipeline for Uber Ride Booking Prediction
Trains a Decision Tree Classifier to predict whether a ride booking is successful (YES) or unsuccessful (NO).
"""

import os
import sys
import time
import joblib
import pandas as pd
import numpy as np
from sklearn.model_selection import train_test_split
from sklearn.tree import DecisionTreeClassifier
from sklearn.preprocessing import LabelEncoder
from sklearn.metrics import (
    accuracy_score,
    precision_score,
    recall_score,
    f1_score,
    confusion_matrix,
    classification_report
)

# Robust path handling
SRC_DIR = os.path.dirname(os.path.abspath(__file__))
BASE_DIR = os.path.dirname(SRC_DIR)
DATA_PATH = os.path.join(BASE_DIR, "data", "ncr_ride_bookings.csv")
MODEL_PATH = os.path.join(BASE_DIR, "models", "decision_tree_model.pkl")

# ML Parameters
TEST_SIZE = 0.20
RANDOM_STATE = 42
DT_MAX_DEPTH = 8
DT_MIN_SAMPLES_LEAF = 20

INPUT_FEATURES = [
    'Vehicle_Type_Enc',
    'Pickup_Location_Enc',
    'Drop_Location_Enc',
    'Avg_VTAT',
    'Avg_CTAT',
    'Booking_Value',
    'Ride_Distance',
    'Hour',
    'DayOfWeek',
    'Month'
]

FEATURE_LABELS = {
    'Vehicle_Type_Enc': 'Vehicle Type',
    'Pickup_Location_Enc': 'Pickup Location',
    'Drop_Location_Enc': 'Drop Location',
    'Avg_VTAT': 'Avg VTAT (min)',
    'Avg_CTAT': 'Avg CTAT (min)',
    'Booking_Value': 'Booking Value (INR)',
    'Ride_Distance': 'Ride Distance (km)',
    'Hour': 'Hour of Day',
    'DayOfWeek': 'Day of Week',
    'Month': 'Month'
}

def load_and_preprocess_data(csv_path=DATA_PATH):
    print(f"Loading dataset from: {csv_path}...")
    df = pd.read_csv(csv_path)
    print(f"Dataset loaded: {len(df):,} records, {df.shape[1]} columns.")

    # 1. Target Definition: Successful = YES (Completed), All other = NO
    df['Actual_Status'] = df['Booking Status'].astype(str).str.strip()
    df['Actual_YES_NO'] = df['Actual_Status'].apply(lambda x: 'YES' if x.lower() in ['completed', 'successful'] else 'NO')
    df['Target'] = (df['Actual_YES_NO'] == 'YES').astype(int)

    # 2. Extract Temporal Features
    df['Date_Parsed'] = pd.to_datetime(df['Date'])
    df['Hour'] = pd.to_datetime(df['Time'], format='%H:%M:%S').dt.hour
    df['DayOfWeek'] = df['Date_Parsed'].dt.day_of_week
    df['Month'] = df['Date_Parsed'].dt.month
    df['DayName'] = df['Date_Parsed'].dt.day_name()

    # 3. Handle Missing Values
    median_vtat = df['Avg VTAT'].median()
    median_ctat = df['Avg CTAT'].median()
    df['Avg_VTAT'] = df['Avg VTAT'].fillna(median_vtat)
    df['Avg_CTAT'] = df['Avg CTAT'].fillna(median_ctat)

    df['Booking_Value'] = df.groupby('Vehicle Type')['Booking Value'].transform(lambda s: s.fillna(s.median()))
    df['Booking_Value'] = df['Booking_Value'].fillna(df['Booking Value'].median())

    df['Ride_Distance'] = df.groupby('Vehicle Type')['Ride Distance'].transform(lambda s: s.fillna(s.median()))
    df['Ride_Distance'] = df['Ride_Distance'].fillna(df['Ride Distance'].median())

    # 4. Categorical Encoding
    encoders = {}
    for cat_col, enc_col in [
        ('Vehicle Type', 'Vehicle_Type_Enc'),
        ('Pickup Location', 'Pickup_Location_Enc'),
        ('Drop Location', 'Drop_Location_Enc')
    ]:
        le = LabelEncoder()
        df[enc_col] = le.fit_transform(df[cat_col].astype(str))
        encoders[cat_col] = le

    X = df[INPUT_FEATURES].copy()
    y = df['Target'].copy()

    return df, X, y, encoders

def train_and_save_model():
    t0 = time.time()
    df, X, y, encoders = load_and_preprocess_data(DATA_PATH)

    print(f"Splitting data into 80% Train ({int(len(X)*0.8):,}) and 20% Test ({int(len(X)*0.2):,})...")
    indices = df.index.to_numpy()
    X_train, X_test, y_train, y_test, idx_train, idx_test = train_test_split(
        X, y, indices,
        test_size=TEST_SIZE,
        random_state=RANDOM_STATE,
        stratify=y
    )

    print(f"Training DecisionTreeClassifier (max_depth={DT_MAX_DEPTH}, min_samples_leaf={DT_MIN_SAMPLES_LEAF})...")
    dt = DecisionTreeClassifier(
        max_depth=DT_MAX_DEPTH,
        min_samples_leaf=DT_MIN_SAMPLES_LEAF,
        random_state=RANDOM_STATE
    )
    dt.fit(X_train, y_train)

    # Predictions & Metrics
    y_pred = dt.predict(X_test)
    y_prob = dt.predict_proba(X_test)[:, 1]

    acc = accuracy_score(y_test, y_pred)
    prec = precision_score(y_test, y_pred)
    rec = recall_score(y_test, y_pred)
    f1 = f1_score(y_test, y_pred)
    cm = confusion_matrix(y_test, y_pred)

    print("\n" + "=" * 55)
    print("           MODEL EVALUATION RESULTS")
    print("=" * 55)
    print(f"Accuracy:  {acc:.4%}")
    print(f"Precision: {prec:.4%}")
    print(f"Recall:    {rec:.4%}")
    print(f"F1 Score:  {f1:.4%}")
    print("\nConfusion Matrix [ [TN, FP], [FN, TP] ]:")
    print(f"  TN={cm[0,0]:,}  FP={cm[0,1]:,}")
    print(f"  FN={cm[1,0]:,}   TP={cm[1,1]:,}")
    print("=" * 55)

    # Save Model
    os.makedirs(os.path.dirname(MODEL_PATH), exist_ok=True)
    joblib.dump(dt, MODEL_PATH)
    print(f"\nTrained model successfully saved to: {MODEL_PATH}")
    print(f"Training completed in {time.time()-t0:.2f} seconds.")

if __name__ == "__main__":
    train_and_save_model()
