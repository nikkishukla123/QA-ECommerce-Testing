# Bug Report

| Bug ID | Module | Title | Severity | Priority | Status |
|---|---|---|---|---|---|
| BUG-001 | Registration | Password mismatch error is not displayed | Medium | High | Open |
| BUG-002 | Login | User can login with leading/trailing spaces in email | Low | Medium | Open |
| BUG-003 | Search | Search result count does not match displayed products | Medium | Medium | Open |
| BUG-004 | Product Details | Product image is broken for one listed item | Medium | High | In Progress |
| BUG-005 | Cart | Quantity update does not refresh total price immediately | High | High | Open |
| BUG-006 | Cart | Removed item reappears after page refresh | High | Critical | Open |
| BUG-007 | Checkout | Invalid PIN code is accepted | High | High | Open |
| BUG-008 | Checkout | Place Order button remains enabled with empty address | Critical | Critical | Open |
| BUG-009 | Payment | Card CVV accepts alphabetic characters | High | High | Open |
| BUG-010 | Order Confirmation | Confirmation page does not display order ID | Medium | High | Open |

## Sample Detailed Bug

### BUG-008: Place Order button remains enabled with empty address
**Module:** Checkout  
**Severity:** Critical  
**Priority:** Critical  
**Status:** Open  

**Steps to Reproduce:**
1. Login with a valid customer account.
2. Add any product to cart.
3. Proceed to checkout.
4. Keep address fields empty.
5. Observe the Place Order button.

**Expected Result:**  
Place Order button should remain disabled until all mandatory address fields are filled.

**Actual Result:**  
Place Order button is enabled even when mandatory address fields are empty.

**Impact:**  
User can attempt to place an order without required shipping details, causing order processing failure.
