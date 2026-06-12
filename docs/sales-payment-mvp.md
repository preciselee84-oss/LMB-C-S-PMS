# Sales Payment Automation MVP

## Scope

This MVP implements the first usable loop for SMB sales and payment operations.

- Sales staff register a claimed customer lead from the web dashboard.
- The backend prevents duplicate customer claims with a unique customer name.
- A bank transaction can be submitted manually while the finance scraper is not connected.
- The matching engine compares depositor name and amount against open leads.
- Exact amount and VAT-included amount are treated as successful matches.
- The dashboard shows lead counts, waiting payments, paid payments, confirmed amount, recent matches, and unpaid risk.

## API Surface

- `POST /api/v1/sales/leads`
- `GET /api/v1/sales/leads`
- `POST /api/v1/sales/transactions/match`
- `GET /api/v1/sales/matches`
- `GET /api/v1/sales/summary`

## Matching Rules

1. Normalize customer name and depositor name by lowercasing and removing spaces.
2. Match when either normalized name contains the other.
3. Match amount when it is exactly the expected amount.
4. Match amount when it is `expected_amount * 1.1`, for VAT-included deposits.

## Next Integrations

- Replace manual `transactions/match` calls with the finance API or scraper polling job.
- Add APScheduler job orchestration and transaction source adapters.
- Add notification dispatch after `PaymentMatch` creation.
- Store notification delivery results separately from raw bank transaction data.
