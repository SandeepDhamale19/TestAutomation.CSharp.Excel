Feature: ExcelMacro
	In order to avoid silly mistakes
	As a math idiot
	I want to be told the sum of two numbers

@Unit
Scenario: Run Excel macro
	Given I have macro enabled workbook
	Then I can run macro