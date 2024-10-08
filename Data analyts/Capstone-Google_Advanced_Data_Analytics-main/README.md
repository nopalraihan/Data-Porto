# Capstone Project for Google Advanced Data Analytics Professional Certificate
# Salifort Employee Retention Project

## Project Overview
The intended audience for this project is the Salifort Human Resources & Leadership Team.  In this scenario, Human Resources has requested that a data analyst review the results from an employee satisfaction survey to develop a model that predicts employees will leave the company.  They have also requested an executive summary.

For this task, I used XGBoost binary classification as an excellent choice in turnover prediction for improved employee retention and reduced cost of re-training new hires.  



## Table of Contents
- [Data Sources](#data-sources)
- [Tools](#tools)
- [Data Cleaning](#data-cleaning) 
- [Exploratory Data Analysis](#exploratory-data-analysis)
- [Data Preparation for Modeling](#data-preparation-for-modeling)
- [Modeling: XGBoost Binomial Classification](#modeling-xgboost-binomial-classification)
  - [Model Results on Training Data](#model-results-on-training-data)
  - [XGBoost Predict on Validation Data](#xgboost-predict-on-validation-data)
  - [XGBoost Predict on Test Data](#xgboost-predict-on-test-data)
- [Conclusion and Insights](#conclusion-and-insights)
- [Recommendation and Next Steps](#recommendation-and-next-steps)
  
### Data Sources
- Google/Coursera provided a Jupyter Notebook with the CSV file of the Salifort employee data which was obtained from [Kaggle.](https://www.kaggle.com/datasets/mfaisalqureshi/hr-analytics-and-job-prediction?select=HR_comma_sep.csv)
  
### Tools
- **Python**: For scripting and running the data analysis.
- **Pandas**: Used for data manipulation and analysis.
- **NumPy**: Employed for numerical computing.
- **Matplotlib/Seaborn**: For creating visualizations.
- **Scikit-learn**: Utilized for machine learning model building.
- **XGBoost**: For XGBClassifier and Feature Importance Plot

### Data Overview
- The data file has 15000 rows (including a header) and 10 columns.
- Column 'Left' is the target, while the other 9 columns are features.

![columns_description](https://github.com/israelh88/Capstone-Google_Advanced_Data_Analytics/blob/852294b8c9eb1ffc406ab781423ea0beb2029259/images/Screenshot%202024-03-03%20155159.png)

### Data Cleaning
- Executed **.info()** to gather basic information about the data
- Executed **.describe()** for descriptive statistics
- Renamed columns and corrected spelling, changed names to snake_case, and made them more concise as needed.
  e.g., 'average_montly_hours' [sic] was changed to 'monthly_hours.'
- Checked for missing values with .isna().sum. No missing values were found.
- Removed duplicates with the method, **.drop_duplicates().**

### Exploratory Data Analysis
- Examined counts of employees who left vs those who stayed.
  - In this data set, 16.6% of employees left, 83.4% stayed.
  - This ratio is approximately the same when grouped by Department.
  
### Data Preparation for Modeling
- Converted categorical columns into numeric columns.
  - Changed Salary from low, medium, high to 0, 1, 2.
  - Converted Department names into numeric values with pd.get_dummies.

### Modeling
- Column 'left' is the Target, **y.**  "Left" as in the employee left the company.
- All other columns are features and were set as **X.**
- I split the data into training/validate/test splits of 60/20/20.
  - X_tr, X_test, y_tr, y_test and set test_size = 0.20.
  - Followed by X_train, X_val, y_train, y_val with test_size = 0.25.
- I set tratify = y (yes) due to the imbalanced data.
- A validation set is used to evaluate the model’s performance.
  
### Modeling: XGBoost Binomial Classification
- Instantiated the XGBoost classifier and set objective='binary:logistic'. 
- Tuned the following hyperparameters (cv_params). 
  - 'max_depth': [8,None], 
  - 'min_child_weight': [1,5],
  - 'learning_rate': [0.2,0.3,0.4],
  - 'n_estimators': [50, 100]
- Instantiated the GridSearchCV object xgb_cv. 
- Used refit = 'recall' to minimize false negatives--predicting an employee stayed when they left.

### Model Results on Training Data
- xgb_cv.fit(X_train, y_train)
- Best score with xgb_cv.best_score_ = 0.9121
- Best parameter settings with xgb_cv.best_params_
  - 'learning_rate': 0.3,
  - 'max_depth': None,
  - 'min_child_weight': 5,
  - 'n_estimators': 50
 
### XGBoost Predict on Validation Data
  - XGB_val_preds = xgb_cv.best_estimator_.predict(X_val)
  - Recall Score = 0.9171

### XGBoost Predict on Test Data
  - xgb_test_preds = xgb_cv.best_estimator_.predict(X_test)
  - Recall Score = 0.9271
 
## Conclusion and Insights
XGBoost classifier is an excellent model for predicting employee turnover, with the champion **recall score** of **0.9271.** This is illustrated in the Confusion Matrix--the model predicted incorrectly that 29 employees stayed, while 369 were correctly labeled as leaving. 
### Confusion Matrix
!["SalifortCapstoneProjectConfusionMatrix"](https://github.com/israelh88/Capstone-Google_Advanced_Data_Analytics/blob/main/images/Screenshot%202024-02-28%20020321.png?raw=true)

### Feature Importance Plot
The Feature Importance Plot is a bar chart that helps us understand which features (variables) have the most significant impact on the target variable (in this case, the likelihood of an employee leaving the company). I've shared below the plot for the XGBoost model results.

!["SalifortCapstoneProjectFeatureImportance](https://github.com/israelh88/Capstone-Google_Advanced_Data_Analytics/blob/main/images/Screenshot%202024-02-28%20154746.png?raw=true)

Although this Pareto is useful, it does not indicate the directionality of the effects of the features. So, my curiosity led me to a Google search on Feature Importance which resulted in learning about **"Partial Dependence Plots,"** a linear plot that goes hand-in-hand with Feature Importance Plots. Although partial dependence was beyond the course’s scope, I gained insight into it independently, inspired by the course’s advice to seek knowledge outside the curriculum.  It is useful in understanding the effects of the features; therefore, I have included the Partial Dependence Plots for the Top 5 features.

### Partial Dependence Plot: Employee Satisfaction
Employee satisfaction predicts turnover, as expected. Satisfaction of 0.1 or less is a definite walk.

![PDPsatisfaction](https://github.com/israelh88/Capstone-Google_Advanced_Data_Analytics/blob/5d8fbb66b05f6ef3268e43f62b82052c94af51dd/images/Screenshot%202024-03-03%20224048.png)

### Partial Dependence Plot: Monthly Hours Worked
Here, the Partial Dependence Plot is quite revealing, as it shows the relationship of the target vs the continuous variable, monthly hours worked. In the plot below, **both overworked and underworked employees tend to leave the company.**
  - Salifort employees are most likely to leave when averaging 250 hours per month, which equates to approximately 63 hours per week.
  - Additionally, employees tend to leave when working 135 hours per month--around 30 hours per week.

![SalifortCapstoneProjectFeatureImportance](https://github.com/israelh88/Capstone-Google_Advanced_Data_Analytics/blob/d82f2c4f8fe969ccecd0af26cf990130c648c749/images/Screenshot%202024-03-22%20053716.png)

NOTE: A standard workweek of 40 hours equals approximately 173 hrs per month.

### Partial Dependence Plot: Last Evaluation
Employees with very low last evaluation scores are more likely to leave the company. Surprisingly, employees with very high last evaluation scores also exhibit a higher likelihood of leaving (target value close to 1). Employees with mid-range evaluation scores tend to stay (target value closer to 0).

![PDP_last eval](https://github.com/israelh88/Capstone-Google_Advanced_Data_Analytics/blob/aee7bb1c4d739b30d7b3d513e85cc26d9ca8da0e/images/Screenshot%202024-03-22%20012713.png)

### Partial Dependence Plot: Number of Projects
It is not unexpected that the number of projects affects employee turnover. At 6 or more projects, there is a tendency for employees to leave the company.  More projects can lead to more hours worked.

![PDPnumber_projects](https://github.com/israelh88/Capstone-Google_Advanced_Data_Analytics/blob/285b43df39bfdfd21b320a18928662a63f378873/images/Screenshot%202024-03-20%20115348.png)  

### Partial Dependence Plot: Tenure
As for tenure, there's something unusual about the higher tendency to leave the company after 5 years. Employees often wait until stock options fully vest before leaving the company.
  
![PDP_tenure](https://github.com/israelh88/Capstone-Google_Advanced_Data_Analytics/blob/628f775d64330ed3f478868f35b97d44795abfab/images/Screenshot%202024-03-20%20111823.png)

### Recommendation and Next Steps
- Alleviate the load on those with 6 or more projects by dividing the work among those with much shorter than 40-hour workweeks. 
- Ascertain why employees leave after 5 years of tenure.
- For employees with excellent evaluations, offer incentives to retain these talented individuals.
- As for last_evaluation, find out why better-performing employees leave.  Do they find a better opportunity elsewhere? Consider offering incentives to longer-term employees.
