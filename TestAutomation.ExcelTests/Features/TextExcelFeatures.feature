Feature: TextExcelFeatures
	In order to avoid silly mistakes
	As a math idiot
	I want to be told the sum of two numbers

@Unit
Scenario: Get active worksheets names
	Given I have valid workbook
	Then I should get active workheet name


@Unit
Scenario: Get All worksheets names
	Given I have valid workbook
	Then I should get all workheet names

@Unit
Scenario: Activate worksheets by name/ index
	Given I have valid workbook
	Then I can activate worksheets by name/ index

@Unit
Scenario: Read excel cell values
	Given I have valid workbook
	Then I can read excel cell values of different format

@Unit
Scenario: Read excel range co-ordinates
	Given I have valid workbook
	Then I can get excel range co-ordinates

@Unit
Scenario: Read excel range as list
	Given I have valid workbook
	Then I can get excel range as list

@Unit
Scenario: Get excel cell background color
	Given I have valid workbook
	Then I can get excel cell background color

@Unit
Scenario: Get all excel cell properties
	Given I have valid workbook
	Then I can get all excel cell properties

@Unit
Scenario: Get all excel range properties
	Given I have valid workbook
	Then I can get all excel range properties

@Unit
Scenario: Get all excel cell formula
	Given I have valid workbook
	Then I can get all excel cell formula

@Unit
Scenario: Get excel cell font color
	Given I have valid workbook
	Then I can get all excel cell font color

@Unit
Scenario: Get excel cell font bold property
	Given I have valid workbook
	Then I can get all excel cell bold property

	@Unit
Scenario: Set excel cell border
	Given I have valid workbook
	Then I can set excel cell border

@Unit
Scenario: Write to excel
	Given I have valid workbook
	Then I can write to excel

@Unit
Scenario: Read chart values
	Given I have valid workbook
	Then I can read chart values