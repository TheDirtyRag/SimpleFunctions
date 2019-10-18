# SimpleFunctions
Bloomberg Functions Simplified


This Add-in is made to simplify and save time from looking up the Bloomberg mnemonics and instead creating functions that directly references the values.

Right now Simple Functions includes major Bank Ratios, Insurance Metrics, and some general metrics. 

To receive the data, you will also need the Bloomberg Excel Add-in as this Add-in is a depedent on it for retrieving the data. 

However, there is a intergrated backup program in the Add-in that allows for error wrapping with the correct values for a given function. 

This means that you only need to retrieve the values once, after that, the values are saved to the last worksheet created. So make sure to create a blank Worksheet before you run the program, call it "Backup" or such. 

The Simple Functions are then wrapped with a IFERROR statement that preserves the values within your tables.


To install
1. Download the latest release.
2. Activate the Developer tab in the ribbon 
3. Click on Excel Add-Ins (The one with the gears)
4. Browse to the folder that contains your add-in
5. Select it from the list by hitting checkbox the left.

To activate the backup program you have 2 options.
1. Press ALT-F8
2. In the input box type "backupvalues" (without quotes) and press run

OR

1. Right click on the ribbon (The menu bar)
2. Select customize the ribbon
3. From the choose commands drop-down menu select MACROS
4. Create a new group (call it whatever you want)
5. Click on backup-values and then press the ADD>> button
6. Click OK
7. Run the backupvalues button.

Some key steps to allow this backup program to work.

1. Always save before you run any program. You cannot undo Macros. 
2. Make sure that your tables are well formatted throughout your spreadsheets. If data starts on cell "A1" then makesure that all your data starts on cell "A1" or it might be missed. 

Further Documentation on the Functions

ClaimsSettlementRatio(Equity, Year) 
Returns the Claims Settlement Ratio for the year 

CombinedRatio(Equity, Year) 
Returns the Combined Ratio for the year

CurrentLiquidity(Equity, Year) 
Returns the Current Liquidity for the year 

InsuranceMargin(Equity, Year) 
Returns the Insurance Margin for the year 

InterestIncomeRatio(Equity, Year) 
Returns the Interest Income Ratio for the year 

InterestSpread(Equity, Year) 
Returns the Interest Spread for the year

InvestmentYield(Equity, Year) 
Returns the Investment Yield for the year 

LiquidAssetsToReserves(Equity, Year) 
Returns the Liquid Assets to Reserves ratio for the year 

LoanToDepositRatio(Equity, Year) 
Returns the Loan To Deposit Ratio for the year 

Net_Income(Equity, Year) 
Returns the Net Income for the selected year 

NetChargeoffRatio(Equity, Year) 
Returns the Net Charge Off Ratio for the year 

NetInterestMargin(Equity, Year) 
Returns the net interest margin for the year 

NonInterestIncomeRatio(Equity, Year) 
Returns the Non-Interest Income Ratio for the year 

OverheadRatio(Equity, Year) 
Returns the Overhead Ratio for the year 

PremiumGrowth(Equity, Year) 
Returns the year over year Premium Growth  

Profit_Margin(Equity, Year)  
Returns the Profit-Margin for the selected year 

ReturnOnAssets(Equity, Year) 
Returns the Return on Assets for the year 

ReturnOnEquity(Equity, Year) 
Returns the Return on Equity for the year 

SolvencyRatio(Equity, Year) 
Returns the Solvency Ratio for the year 

TierOneCapital(Equity, Year) 
Returns the Tier One capital for the year 

TierOneCapitalRatio(Equity, Year) 
Returns the Tier One Capital Ratio for the year 

TierOneLeverageRatio(Equity, Year) 
Returns the Tier One Leverage Ratio for the year 
