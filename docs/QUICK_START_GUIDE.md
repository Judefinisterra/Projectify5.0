# Projectify5.0 Quick Start Guide

## Get Started in 10 Minutes

This guide will help you create your first financial model with Projectify5.0 in just 10 minutes.

## Prerequisites

- Microsoft Excel 2016 or later
- Internet connection
- Basic understanding of financial modeling concepts

## Step 1: Install the Add-in (2 minutes)

1. **Download the Add-in**
   - Get the latest release from your organization's IT department
   - Or clone from the repository if you have access

2. **Sideload in Excel**
   - Open Excel
   - Go to `Insert` > `My Add-ins` > `Upload My Add-in`
   - Select `manifest.xml` from the project folder
   - Click `OK`

3. **Verify Installation**
   - Look for "Build Financial Models with AI" in the Excel ribbon
   - Click the button to open the task pane

## Step 2: Choose Your Mode (30 seconds)

When you first open Projectify5.0, you'll see two options:

- **Client Mode**: Simplified interface for quick model creation
- **Developer Mode**: Advanced features for power users

👉 **For this guide, select Client Mode**

## Step 3: Create Your First Model (5 minutes)

### Example: Simple Restaurant Revenue Model

1. **Describe Your Business**
   Type this into the chat interface:
   ```
   I have a pizza restaurant. We sell 100 pizzas per day at $15 each. 
   Our food costs are 30% of revenue. We have 3 employees earning $3,000 per month each.
   Rent is $2,000 per month.
   ```

2. **Review the Model Plan**
   The AI will generate a structured plan like:
   ```
   Revenue and Direct Costs: Pizza sales with 100 units daily at $15 each, 30% food costs
   Corporate Overhead: 3 employees at $3,000/month each, $2,000/month rent
   ```

3. **Confirm and Generate**
   - Type "yes" or "proceed" to confirm
   - The system will generate the Excel model automatically

## Step 4: Review Your Model (2 minutes)

After generation, you'll see:

### New Worksheets Created:
- **Revenue and Direct Costs**: Your pizza sales calculations
- **Corporate Overhead**: Employee and rent expenses
- **Financials**: Complete 3-statement financial model

### Key Features to Notice:
- **Monthly columns**: See your finances month by month
- **Annual totals**: Year-end summaries in the rightmost columns
- **Formulas**: All calculations are linked and dynamic
- **Financial statements**: Income Statement, Balance Sheet, Cash Flow

## Step 5: Customize Your Model (1 minute)

### Change Assumptions:
- Go to the assumption worksheets
- Modify any blue-highlighted cells
- Watch the financial statements update automatically

### Example Changes:
- Change pizza price from $15 to $18
- Increase daily sales from 100 to 120
- Adjust employee salaries

## Common Use Cases

### SaaS Business Model
```
I have a SaaS business with $50 monthly subscriptions. 
I currently have 1,000 customers with 5% monthly churn.
I acquire 100 new customers per month at $200 per customer acquisition cost.
```

### E-commerce Business
```
I sell widgets online for $25 each with $10 cost of goods sold.
I sell 500 units per month and expect 20% monthly growth.
I have $5,000 monthly marketing expenses and $2,000 monthly overhead.
```

### Consulting Business
```
I'm a consultant billing $150 per hour. I work 40 hours per week.
I have $1,000 monthly office expenses and $500 monthly software costs.
```

## Tips for Success

### Writing Good Prompts
✅ **Good**: "I sell 100 units at $10 each with 20% margins"
❌ **Avoid**: "I have a business that makes money"

### Include These Details:
- **Volume**: How many units/customers/transactions
- **Pricing**: Price per unit or subscription fee
- **Costs**: Direct costs, employee costs, overhead
- **Growth**: Expected growth rates or trends
- **Timing**: When expenses occur (monthly, annually)

### Common Patterns:
- **Unit × Price = Revenue**
- **Revenue × Margin % = Gross Profit**
- **# Employees × Salary = Payroll Cost**

## Troubleshooting

### "The AI didn't understand my request"
- Be more specific about numbers and timing
- Break complex requests into simpler parts
- Use the examples above as templates

### "The model looks wrong"
- Check the assumption sheets for blue-highlighted cells
- Verify that your inputs match what you intended
- Use the validation features in Developer Mode

### "I need to add more complexity"
- Switch to Developer Mode for advanced features
- Use the training data system to improve AI understanding
- Consult the full documentation for advanced use cases

## Next Steps

### Learn More:
- Read the full [README.md](../README.md) for comprehensive documentation
- Check out the [Codestring Reference Guide](CODESTRING_REFERENCE.md)
- Explore [Advanced Examples](ADVANCED_EXAMPLES.md)

### Get Help:
- Join the user community
- Contact support team
- Submit feature requests

### Contribute:
- Report bugs and issues
- Share your model templates
- Help improve the AI training data

## Example Models Gallery

### Subscription Business
- Monthly recurring revenue model
- Churn and retention calculations
- Customer acquisition cost analysis

### Manufacturing Business
- Production capacity planning
- Inventory management
- Supply chain cost modeling

### Real Estate Investment
- Property cash flow analysis
- Depreciation calculations
- Financing scenarios

### Startup Fundraising
- Equity dilution modeling
- Investor return calculations
- Exit scenario planning

---

**Need help?** The Projectify5.0 team is here to support you. Check our documentation or contact support for assistance with your financial modeling needs.

**Ready for more?** Switch to Developer Mode to access advanced features like custom codestrings, training data management, and model validation tools. 