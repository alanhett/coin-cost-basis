# Coin Cost Basis

This macro-enabled spreadsheet will help you track cost basis and long-term or short-term treatment for your cryptocurrency trades. It uses the first-in, first-out (FIFO) cost method, which is commonly used for tax compliance.

## Usage

- Download and open coin-cost-basis.xlsm
- Enable macros if prompted
- Create a copy of the blank sheet for each coin or asset you trade
- Enter your buys in the left-hand buy columns. Include any fees related to your purchase
- Enter your sales in the right-hand sell columns. Net out any fees related to your sale
- Each transaction must include a date
- The notes column is optional, and a great place to track fees, or trades between different assets
- Add rows as needed

## Advanced Usage
If you want to add columns, or change the layout, you'll need to open the VBA module and update the constant variables, ensuring that they are mapped to the correct columns.

## Changelog

### 0.5
- Rows are now unlimited. Add rows to the template as needed.
- Additional date validation added.

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

*Disclaimer: This spreadsheet does not constitute legal or tax advice.  Tax laws and regulations change frequently, and their application can vary widely based on the specific facts and circumstances involved. You are responsible for consulting with your own professional tax advisors concerning specific tax circumstances for your business. I disclaim any responsibility for the accuracy or adequacy of any positions taken by you in your tax returns.*