🚗 Uber Ride Booking Prediction & Analytics Dashboard

A Machine Learning project that uses the **NCR Ride Bookings dataset** to predict whether a ride booking will be successful or unsuccessful using a **Decision Tree Classifier**. The model predictions and ride analytics are presented through an interactive **Microsoft Excel Dashboard**.

## 📌 Project Overview

This project combines **Machine Learning, Data Analysis, and Excel Visualization** to analyze Uber ride booking data and build a predictive system.

The system:

- Cleans and preprocesses ride booking data
- Performs feature selection and preprocessing
- Trains a Decision Tree Classification model
- Predicts successful and unsuccessful ride bookings
- Evaluates model performance
- Generates prediction results
- Creates an Excel-based analytics dashboard

## 🎯 Objective

The main objective is to develop a machine learning model that can classify ride bookings as:

- **YES** → Successful booking
- **NO** → Unsuccessful booking

The predictions are then visualized through an Excel dashboard for easier analysis and presentation.

## 🤖 Machine Learning Model

**Algorithm:** Decision Tree Classifier

### Target Variable

`Booking Status`

The target is converted into a binary classification:

| Booking Status        | Prediction |
| --------------------- | ---------- |
| Completed             | YES        |
| Cancelled by Customer | NO         |
| Cancelled by Driver   | NO         |
| Incomplete            | NO         |

### Dataset

- **Dataset:** NCR Ride Bookings
- **Records:** 150,000
- **Features:** Ride booking and operational attributes
- **Training Data:** 120,000 records
- **Testing Data:** 30,000 records

## 📊 Model Performance

| Metric    |      Score |
| --------- | ---------: |
| Accuracy  | **95.55%** |
| Precision | **93.34%** |
| Recall    | **99.96%** |
| F1 Score  | **96.54%** |

### Confusion Matrix

- True Negative: 10,074
- False Positive: 1,326
- False Negative: 8
- True Positive: 18,592

## 📈 Excel Dashboard

The project includes a professional Excel dashboard containing:

### KPI Cards

- Total Bookings
- Successful Bookings
- Unsuccessful Bookings
- Success Rate
- Model Accuracy
- Average Ride Distance
- Average Booking Value

### Visualizations

- Booking Success Status
- Bookings by Vehicle Type
- Actual vs Predicted Results
- Hourly Booking Demand
- Average Booking Value by Vehicle
- Average Ride Distance by Vehicle
- Top 10 Pickup Locations

## 📁 Project Structure

```text
Uber-Ride-ML-Dashboard/
│
├── data/
│   └── ncr_ride_bookings.csv
│
├── models/
│   └── decision_tree_model.pkl
│
├── output/
│   └── Uber_Ride_ML_Dashboard_Final.xlsx
│
├── src/
│   ├── train_model.py
│   └── create_dashboard.py
│
├── requirements.txt
└── README.md
```

## 🛠️ Technologies Used

- **Python**
- **Pandas**
- **NumPy**
- **Scikit-learn**
- **Matplotlib**
- **OpenPyXL**
- **Microsoft Excel**
- **GitHub**

## ⚙️ Project Workflow

```text
NCR Ride Booking Dataset
          ↓
    Data Cleaning
          ↓
 Feature Engineering
          ↓
   Feature Selection
          ↓
   Train/Test Split
          ↓
 Decision Tree Classifier
          ↓
     Prediction
          ↓
   Model Evaluation
          ↓
    Excel Dashboard
```

## 🚀 How to Run

### 1. Clone the repository

```bash
git clone https://github.com/rashmichaudhari902/Uber-Ride-ML-Dashboard.git
cd Uber-Ride-ML-Dashboard
```

### 2. Install dependencies

```bash
pip install -r requirements.txt
```

### 3. Run the model

```bash
python src/train_model.py
```

### 4. Generate the dashboard

```bash
python src/create_dashboard.py
```

The generated Excel dashboard will be available in the `output` folder.

## 📂 Project Outputs

The project generates:

- Trained Decision Tree model
- Ride booking predictions
- Model evaluation metrics
- Feature importance
- Excel analytics dashboard

## 🔍 Key Insights

The project provides insights into:

- Ride booking success rate
- Vehicle-wise booking patterns
- Hourly ride demand
- Pickup location demand
- Average booking value
- Average ride distance
- Model prediction performance

## 👩‍💻 Project

**Uber Ride Booking Prediction & Analytics Dashboard**

Built using Machine Learning and Excel for predictive analytics and visualization.
