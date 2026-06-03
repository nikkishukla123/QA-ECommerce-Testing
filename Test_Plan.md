# Test Plan: E-Commerce Website Testing

## Objective
To verify that the e-commerce website works correctly for customer-facing workflows such as registration, login, product search, cart management, checkout, and order confirmation.

## Scope
### In Scope
- Registration with valid and invalid data
- Login, logout, and forgot password
- Product search and filtering
- Product details validation
- Add to cart and cart update
- Checkout address validation
- Payment field validation
- Order confirmation
- Regression testing for critical flows

### Out of Scope
- Real payment gateway processing
- Backend inventory integration
- Performance testing
- Security penetration testing

## Test Environment
- Browser: Chrome, Edge
- Device: Desktop and mobile responsive view
- Operating System: Windows
- Test Data: Valid/invalid email, password, address, product keywords, coupon codes

## Entry Criteria
- Application build is deployed and accessible.
- Major modules are available for testing.
- Test data is prepared.
- Requirements or user stories are available.

## Exit Criteria
- All planned test cases are executed.
- Critical and high severity bugs are fixed or accepted.
- Regression testing is completed.
- Test summary report is prepared.

## Risks
- Payment gateway cannot be fully validated without sandbox integration.
- Requirements may change during testing.
- Some scenarios may depend on product inventory availability.

## Test Deliverables
- Test Cases Workbook
- Bug Report
- Mock JIRA Tickets
- Requirement Traceability Matrix
- Test Summary
