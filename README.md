# Coin Cost Basis

This macro-enabled spreadsheet helps you keep track of the cost basis of your cryptyocurrecny trades. Using the first-in, first-out (FIFO) cost method (the most commonly used cost method for calculating capital gains), Coin Cost Basis will help you calculate, and distinguish between long-term and short-term gains and losses.

## Changelog

### 0.4
- Minor bug fixes as a result of testing
- Updated long-term test to more strictly adhere to guidance (using anniversary date vs 365 days)

### 0.3
- Moved date and coin value validation to its own function to prevent partial calculation
- Added commenting to coin splits

### 0.2
- Sales with long-term and short-term gains are split
- Error handling for coin sales beyond lot totals
- Date validation and data validation added
- Added status column for lot status and short vs. long-term gains and losses

### 0.1
- Built initial logic for FIFO cost basis
  
## Usage

- Create a copy of the blank spreadsheet for each coin you trade.
- Enter your buys in the buys columns. Include the total coins acquired, and the cost in your local currency (make sure to include any fees related to your purchase, these are typically eligible components of cost basis).
- Enter your sales in the sell columns (again, netting out any fees from the amount received).
- Each transaction needs a date, the note is optional.

*Disclaimer: This spreadsheet does not constitute legal or tax advice.  Tax laws and regulations change frequently, and their application can vary widely based on the specific facts and circumstances involved. You are responsible for consulting with your own professional tax advisors concerning specific tax circumstances for your business. Alan Hettinger disclaims any responsibility for the accuracy or adequacy of any positions taken by you in your tax returns.*