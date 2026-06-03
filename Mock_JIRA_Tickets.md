# Mock JIRA Defect Tickets

## EC-101: Cart total does not update after quantity change
**Issue Type:** Bug  
**Module:** Cart  
**Severity:** High  
**Priority:** High  
**Assignee:** QA Tester  
**Status:** Open  

**Description:** When the user changes item quantity in the cart, the item total and cart subtotal are not updated until the page is refreshed.

**Steps to Reproduce:**
1. Login to the application.
2. Search for any available product.
3. Add product to cart.
4. Open cart page.
5. Change quantity from 1 to 2.

**Expected Result:** Cart total should update immediately.

**Actual Result:** Cart total remains unchanged until refresh.

## EC-102: Checkout allows empty address submission
**Issue Type:** Bug  
**Module:** Checkout  
**Severity:** Critical  
**Priority:** Critical  
**Status:** Open  

**Description:** The Place Order button remains enabled even when mandatory address fields are blank.

## EC-103: Search result count mismatch
**Issue Type:** Bug  
**Module:** Search  
**Severity:** Medium  
**Priority:** Medium  
**Status:** Open  

**Description:** Search results page displays a count that does not match the actual number of products shown.

## EC-104: CVV field accepts alphabetic characters
**Issue Type:** Bug  
**Module:** Payment  
**Severity:** High  
**Priority:** High  
**Status:** Open  

**Description:** CVV field accepts non-numeric characters during checkout payment validation.
