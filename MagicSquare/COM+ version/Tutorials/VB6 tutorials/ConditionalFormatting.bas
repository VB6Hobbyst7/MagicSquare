Attribute VB_Name = "ConditionalFormatting"
Public CONDITIONALFORMATTING_CONDITIONAL_FORMATTING_TYPE_COMPARE_CELL_VALUE As Long
Public CONDITIONALFORMATTING_CONDITIONAL_FORMATTING_TYPE_EVALUATE_FORMULA As Long
Public CONDITIONALFORMATTING_OPERATOR_NO_COMPARATION As Long
Public CONDITIONALFORMATTING_OPERATOR_BETWEEN As Long
Public CONDITIONALFORMATTING_OPERATOR_NOT_BETWEEN As Long
Public CONDITIONALFORMATTING_OPERATOR_EQUAL_TO As Long
Public CONDITIONALFORMATTING_OPERATOR_NOT_EQUAL_TO As Long
Public CONDITIONALFORMATTING_OPERATOR_GREATER_THAN As Long
Public CONDITIONALFORMATTING_OPERATOR_LESS_THAN As Long
Public CONDITIONALFORMATTING_OPERATOR_GREATER_THAN_OR_EQUAL_TO As Long
Public CONDITIONALFORMATTING_OPERATOR_LESS_THAN_OR_EQUAL_TO As Long

Sub Initialize()
	CONDITIONALFORMATTING_CONDITIONAL_FORMATTING_TYPE_COMPARE_CELL_VALUE = 1
	CONDITIONALFORMATTING_CONDITIONAL_FORMATTING_TYPE_EVALUATE_FORMULA = 2
	CONDITIONALFORMATTING_OPERATOR_NO_COMPARATION = 0
	CONDITIONALFORMATTING_OPERATOR_BETWEEN = 1
	CONDITIONALFORMATTING_OPERATOR_NOT_BETWEEN = 2
	CONDITIONALFORMATTING_OPERATOR_EQUAL_TO = 3
	CONDITIONALFORMATTING_OPERATOR_NOT_EQUAL_TO = 4
	CONDITIONALFORMATTING_OPERATOR_GREATER_THAN = 5
	CONDITIONALFORMATTING_OPERATOR_LESS_THAN = 6
	CONDITIONALFORMATTING_OPERATOR_GREATER_THAN_OR_EQUAL_TO = 7
	CONDITIONALFORMATTING_OPERATOR_LESS_THAN_OR_EQUAL_TO = 8
End Sub