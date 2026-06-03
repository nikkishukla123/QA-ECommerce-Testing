import fs from "node:fs/promises";
import path from "node:path";
import { fileURLToPath } from "node:url";
import { SpreadsheetFile, Workbook } from "@oai/artifact-tool";

const __dirname = path.dirname(fileURLToPath(import.meta.url));
const workbook = Workbook.create();

const colors = {
  navy: "#1F3A5F",
  teal: "#0F766E",
  white: "#FFFFFF",
  lightBlue: "#E0F2FE",
  grid: "#D1D5DB",
};

const modules = [
  ["Registration", 7],
  ["Login", 7],
  ["Search", 8],
  ["Product Details", 6],
  ["Cart", 11],
  ["Checkout", 8],
  ["Payment", 4],
  ["Order Confirmation", 1],
  ["Regression", 3],
];

const scenarios = {
  Registration: [
    ["Verify registration page opens", "Open application > Click Sign Up", "Registration form should display", "Registration form displayed", "Pass", "High", "Smoke"],
    ["Register with valid user data", "Enter name, valid email, mobile, password > Submit", "Account should be created successfully", "Account created successfully", "Pass", "High", "Functional"],
    ["Register with existing email", "Enter already registered email > Submit", "Duplicate email validation should display", "Validation displayed", "Pass", "Medium", "Negative"],
    ["Register with invalid email format", "Enter abc@test > Submit", "Invalid email message should display", "Invalid email message displayed", "Pass", "Medium", "Negative"],
    ["Submit registration with blank mandatory fields", "Keep mandatory fields blank > Submit", "Required field validation should display", "Validation displayed", "Pass", "High", "Negative"],
    ["Verify password and confirm password mismatch", "Enter different password values > Submit", "Password mismatch message should display", "No mismatch message displayed", "Fail", "High", "Functional"],
    ["Verify mobile number accepts only digits", "Enter alphabetic mobile number > Submit", "Mobile field should reject alphabets", "Alphabets rejected", "Pass", "Medium", "Boundary"],
  ],
  Login: [
    ["Login with valid credentials", "Enter valid email and password > Login", "User should login successfully", "User logged in", "Pass", "High", "Smoke"],
    ["Login with invalid password", "Enter valid email and wrong password > Login", "Invalid credentials message should display", "Message displayed", "Pass", "High", "Negative"],
    ["Login with unregistered email", "Enter unregistered email > Login", "User should not login", "Login blocked", "Pass", "Medium", "Negative"],
    ["Login with blank email and password", "Click Login without data", "Required field validation should display", "Validation displayed", "Pass", "High", "Negative"],
    ["Verify forgot password link", "Click Forgot Password", "Forgot password screen should open", "Screen opened", "Pass", "Medium", "Functional"],
    ["Verify logout functionality", "Login > Click Logout", "User should logout and session should end", "Logout successful", "Pass", "High", "Regression"],
    ["Verify email field trims spaces", "Enter email with leading and trailing spaces > Login", "System should trim or validate spaces", "Login allowed without warning", "Fail", "Low", "UI"],
  ],
  Search: [
    ["Search with valid product keyword", "Search for mobile", "Relevant products should display", "Relevant products displayed", "Pass", "High", "Functional"],
    ["Search with invalid keyword", "Search for random text xyz123", "No results message should display", "No results message displayed", "Pass", "Medium", "Negative"],
    ["Search with blank input", "Click search with blank field", "System should show validation or all products as per requirement", "All products displayed", "Pass", "Low", "UI"],
    ["Verify search suggestions", "Type partial keyword in search box", "Suggestions should appear", "Suggestions displayed", "Pass", "Medium", "Functional"],
    ["Verify product count on result page", "Search laptop and compare count", "Displayed count should match visible products", "Count mismatch found", "Fail", "Medium", "Functional"],
    ["Verify filter by price", "Apply price filter", "Products should match selected price range", "Products filtered", "Pass", "High", "Functional"],
    ["Verify sort by price low to high", "Apply low-to-high sorting", "Products should sort ascending by price", "Sorted correctly", "Pass", "Medium", "Regression"],
    ["Verify search after clearing filters", "Apply filters > Clear filters", "Original result list should restore", "Results restored", "Pass", "Medium", "Regression"],
  ],
  "Product Details": [
    ["Open product details page", "Click any product from listing", "Product details page should open", "Page opened", "Pass", "High", "Smoke"],
    ["Verify product image loads", "Open product details page", "Main image should load correctly", "One product image broken", "Fail", "High", "UI"],
    ["Verify product price is visible", "Open product details page", "Price should display clearly", "Price displayed", "Pass", "High", "Functional"],
    ["Verify product rating and reviews", "Open product details page", "Rating and reviews should display", "Displayed", "Pass", "Low", "UI"],
    ["Verify delivery PIN code check", "Enter valid PIN code", "Delivery availability should display", "Availability displayed", "Pass", "Medium", "Functional"],
    ["Verify invalid PIN code message", "Enter invalid PIN code", "Invalid PIN code message should display", "Invalid PIN accepted", "Fail", "High", "Negative"],
  ],
  Cart: [
    ["Add product to cart from product details", "Click Add to Cart", "Product should be added to cart", "Product added", "Pass", "High", "Smoke"],
    ["Add same product twice", "Click Add to Cart twice", "Quantity should increase or duplicate should be handled", "Quantity increased", "Pass", "Medium", "Functional"],
    ["Verify cart icon count", "Add product to cart", "Cart count should update", "Cart count updated", "Pass", "Medium", "Functional"],
    ["Update product quantity in cart", "Change quantity from 1 to 2", "Cart subtotal should update immediately", "Subtotal not updated until refresh", "Fail", "High", "Functional"],
    ["Remove product from cart", "Click Remove on cart item", "Item should be removed", "Item removed", "Pass", "High", "Functional"],
    ["Verify removed product after refresh", "Remove item > Refresh page", "Removed item should not reappear", "Removed item reappeared", "Fail", "Critical", "Regression"],
    ["Verify cart with multiple products", "Add multiple products", "All products should display with correct totals", "Displayed correctly", "Pass", "High", "Functional"],
    ["Verify out of stock product cannot be added", "Open out-of-stock product", "Add to Cart should be disabled", "Button disabled", "Pass", "High", "Functional"],
    ["Apply valid coupon code", "Enter valid coupon > Apply", "Discount should apply", "Discount applied", "Pass", "Medium", "Functional"],
    ["Apply invalid coupon code", "Enter invalid coupon > Apply", "Invalid coupon message should display", "Message displayed", "Pass", "Medium", "Negative"],
    ["Proceed to checkout from cart", "Click Proceed to Checkout", "Checkout page should open", "Checkout opened", "Pass", "High", "Smoke"],
  ],
  Checkout: [
    ["Verify checkout page loads", "Open checkout page", "Address and payment sections should display", "Sections displayed", "Pass", "High", "Smoke"],
    ["Submit checkout with valid address", "Enter valid shipping details", "Address should be accepted", "Address accepted", "Pass", "High", "Functional"],
    ["Submit checkout with blank address", "Keep address blank and continue", "Mandatory validation should display", "Place Order enabled", "Fail", "Critical", "Negative"],
    ["Verify city field accepts text only", "Enter numeric city", "City validation should display", "Validation displayed", "Pass", "Medium", "Negative"],
    ["Verify PIN code length validation", "Enter 3 digit PIN", "PIN length validation should display", "Validation displayed", "Pass", "High", "Boundary"],
    ["Select saved address", "Choose saved address", "Selected address should populate", "Address populated", "Pass", "Medium", "Functional"],
    ["Edit saved address", "Click Edit > Update address", "Updated address should save", "Address saved", "Pass", "Medium", "Regression"],
    ["Verify order summary total", "Compare product total, discount, shipping, final amount", "Final amount should be correct", "Amount correct", "Pass", "High", "Functional"],
  ],
  Payment: [
    ["Pay with valid card details", "Enter valid card number, expiry, CVV", "Payment should process in sandbox", "Payment processed", "Pass", "High", "Functional"],
    ["Pay with invalid card number", "Enter invalid card number", "Invalid card validation should display", "Validation displayed", "Pass", "High", "Negative"],
    ["Verify CVV accepts digits only", "Enter alphabetic CVV", "CVV should reject alphabets", "Alphabetic CVV accepted", "Fail", "High", "Negative"],
    ["Verify expired card validation", "Enter expired card date", "Expired card message should display", "Message displayed", "Pass", "High", "Negative"],
  ],
  "Order Confirmation": [
    ["Verify confirmation after order placement", "Complete checkout", "Order confirmation page should display", "Page displayed but order ID missing", "Fail", "High", "Functional"],
  ],
  Regression: [
    ["Verify login to checkout end-to-end flow", "Login > Search > Add to Cart > Checkout", "Critical flow should complete", "Flow completed with known checkout issue", "Fail", "Critical", "Regression"],
    ["Verify search and cart after bug fix build", "Search product > Add to cart > Update quantity", "Previously working functions should still pass", "Passed except subtotal issue", "Fail", "High", "Regression"],
    ["Verify logout after order attempt", "Attempt order > Logout", "Session should end safely", "Logout successful", "Pass", "Medium", "Regression"],
  ],
};

const testCases = [];
let counter = 1;
for (const [module] of modules) {
  for (const row of scenarios[module]) {
    testCases.push([`TC-${String(counter).padStart(3, "0")}`, module, ...row]);
    counter += 1;
  }
}

const bugs = [
  ["BUG-001", "TC-006", "Registration", "Password mismatch error is not displayed", "Medium", "High", "Open", "Different password and confirm password accepted without clear error"],
  ["BUG-002", "TC-014", "Login", "Email field allows leading/trailing spaces without warning", "Low", "Medium", "Open", "Login works with padded email, validation behavior inconsistent"],
  ["BUG-003", "TC-019", "Search", "Search result count does not match displayed products", "Medium", "Medium", "Open", "Count shows 48 while only 45 products are visible"],
  ["BUG-004", "TC-024", "Product Details", "Product image is broken for one listed item", "Medium", "High", "In Progress", "Broken image icon appears on details page"],
  ["BUG-005", "TC-032", "Cart", "Quantity update does not refresh total price immediately", "High", "High", "Open", "Subtotal updates only after browser refresh"],
  ["BUG-006", "TC-034", "Cart", "Removed item reappears after page refresh", "High", "Critical", "Open", "Item removed from cart appears again after refresh"],
  ["BUG-007", "TC-028", "Product Details", "Invalid PIN code is accepted", "High", "High", "Open", "Invalid PIN code displays delivery available"],
  ["BUG-008", "TC-042", "Checkout", "Place Order button remains enabled with empty address", "Critical", "Critical", "Open", "Mandatory address validation missing before order placement"],
  ["BUG-009", "TC-050", "Payment", "Card CVV accepts alphabetic characters", "High", "High", "Open", "CVV field accepts abc instead of digits only"],
  ["BUG-010", "TC-052", "Order Confirmation", "Confirmation page does not display order ID", "Medium", "High", "Open", "Order success page displays without order reference number"],
];

const rtm = [
  ["REQ-001", "User should be able to register with valid details", "TC-001 to TC-007", "Covered"],
  ["REQ-002", "User should be able to login and logout", "TC-008 to TC-014", "Covered"],
  ["REQ-003", "User should be able to search products", "TC-015 to TC-022", "Covered"],
  ["REQ-004", "User should be able to view product details", "TC-023 to TC-028", "Covered"],
  ["REQ-005", "User should be able to add and update cart items", "TC-029 to TC-039", "Covered"],
  ["REQ-006", "User should be able to checkout with valid address", "TC-040 to TC-047", "Covered"],
  ["REQ-007", "User should receive order confirmation after successful checkout", "TC-048 to TC-052", "Covered"],
  ["REQ-008", "Critical flows should pass regression testing", "TC-053 to TC-055", "Covered"],
];

function title(sheet, range, text) {
  const r = sheet.getRange(range);
  r.merge();
  r.values = [[text]];
  r.format = { fill: colors.navy, font: { bold: true, color: colors.white, size: 16 }, horizontalAlignment: "center" };
}

function header(range) {
  range.format = { fill: colors.teal, font: { bold: true, color: colors.white }, wrapText: true, horizontalAlignment: "center" };
}

function widths(sheet, values) {
  values.forEach((width, idx) => {
    sheet.getRangeByIndexes(0, idx, 1, 1).format.columnWidthPx = width;
  });
}

function border(range) {
  range.format.borders = { all: { style: "continuous", color: colors.grid } };
}

const dashboard = workbook.worksheets.add("Dashboard");
dashboard.showGridLines = false;
title(dashboard, "A1:H1", "E-Commerce Website Testing - QA Summary");
const passCount = testCases.filter((row) => row[6] === "Pass").length;
const failCount = testCases.length - passCount;
dashboard.getRange("A3:B8").values = [
  ["Metric", "Value"],
  ["Total Test Cases", testCases.length],
  ["Passed", passCount],
  ["Failed", failCount],
  ["Bugs Logged", bugs.length],
  ["Pass Rate", passCount / testCases.length],
];
header(dashboard.getRange("A3:B3"));
dashboard.getRange("B8").format.numberFormat = "0.0%";
border(dashboard.getRange("A3:B8"));
dashboard.getRange("A10:F19").values = [
  ["Module", "Total", "Pass", "Fail", "Critical/High Bugs", "Status"],
  ...modules.map(([module]) => [
    module,
    testCases.filter((row) => row[1] === module).length,
    testCases.filter((row) => row[1] === module && row[6] === "Pass").length,
    testCases.filter((row) => row[1] === module && row[6] === "Fail").length,
    bugs.filter((bug) => bug[2] === module && ["Critical", "High"].includes(bug[4])).length,
    testCases.some((row) => row[1] === module && row[6] === "Fail") ? "Needs Fix" : "Passed",
  ]),
];
header(dashboard.getRange("A10:F10"));
border(dashboard.getRange("A10:F19"));
dashboard.getRange("A21:H26").values = [
  ["Project Scope", "", "", "", "", "", "", ""],
  ["Registration, Login, Search, Product Details, Add to Cart, Checkout, Payment, Order Confirmation", "", "", "", "", "", "", ""],
  ["Testing Types", "", "", "", "", "", "", ""],
  ["Functional, Regression, Smoke, Negative, UI, Boundary Value Testing", "", "", "", "", "", "", ""],
  ["Tools", "", "", "", "", "", "", ""],
  ["Excel, JIRA / Mock JIRA, Browser DevTools", "", "", "", "", "", "", ""],
];
for (const row of [21, 22, 23, 24, 25, 26]) dashboard.getRange(`A${row}:H${row}`).merge();
dashboard.getRange("A21:H21").format = { fill: colors.lightBlue, font: { bold: true } };
dashboard.getRange("A23:H23").format = { fill: colors.lightBlue, font: { bold: true } };
dashboard.getRange("A25:H25").format = { fill: colors.lightBlue, font: { bold: true } };
widths(dashboard, [180, 100, 90, 90, 135, 110, 80, 80]);

const tc = workbook.worksheets.add("Test Cases");
title(tc, "A1:I1", "Manual Test Cases");
tc.getRange("A3:I3").values = [["Test Case ID", "Module", "Test Scenario", "Test Steps", "Expected Result", "Actual Result", "Status", "Priority", "Type"]];
tc.getRange(`A4:I${testCases.length + 3}`).values = testCases;
header(tc.getRange("A3:I3"));
border(tc.getRange(`A3:I${testCases.length + 3}`));
tc.getRange(`A4:I${testCases.length + 3}`).format.wrapText = true;
tc.freezePanes.freezeRows(3);
widths(tc, [90, 120, 240, 300, 260, 240, 80, 90, 110]);

const bugSheet = workbook.worksheets.add("Bug Report");
title(bugSheet, "A1:H1", "Defect Log");
bugSheet.getRange("A3:H3").values = [["Bug ID", "Linked TC", "Module", "Bug Title", "Severity", "Priority", "Status", "Description"]];
bugSheet.getRange(`A4:H${bugs.length + 3}`).values = bugs;
header(bugSheet.getRange("A3:H3"));
border(bugSheet.getRange(`A3:H${bugs.length + 3}`));
bugSheet.getRange(`A4:H${bugs.length + 3}`).format.wrapText = true;
bugSheet.freezePanes.freezeRows(3);
widths(bugSheet, [90, 80, 130, 260, 90, 90, 110, 320]);

const jira = workbook.worksheets.add("Mock JIRA");
title(jira, "A1:G1", "Mock JIRA Tickets");
jira.getRange("A3:G3").values = [["JIRA ID", "Bug ID", "Issue Type", "Summary", "Assignee", "Sprint", "Status"]];
jira.getRange(`A4:G${bugs.length + 3}`).values = bugs.map((bug, idx) => [`EC-${101 + idx}`, bug[0], "Bug", bug[3], "QA Tester", "Sprint 1", bug[6]]);
header(jira.getRange("A3:G3"));
border(jira.getRange(`A3:G${bugs.length + 3}`));
jira.getRange(`A4:G${bugs.length + 3}`).format.wrapText = true;
widths(jira, [90, 90, 90, 320, 120, 90, 110]);

const rtmSheet = workbook.worksheets.add("RTM");
title(rtmSheet, "A1:D1", "Requirement Traceability Matrix");
rtmSheet.getRange("A3:D3").values = [["Requirement ID", "Requirement", "Test Case Range", "Status"]];
rtmSheet.getRange(`A4:D${rtm.length + 3}`).values = rtm;
header(rtmSheet.getRange("A3:D3"));
border(rtmSheet.getRange(`A3:D${rtm.length + 3}`));
rtmSheet.getRange(`A4:D${rtm.length + 3}`).format.wrapText = true;
widths(rtmSheet, [110, 420, 160, 100]);

const resume = workbook.worksheets.add("Resume Points");
title(resume, "A1:D1", "Resume Ready Content");
resume.getRange("A3:D11").values = [
  ["Project Title", "E-Commerce Website Testing", "", ""],
  ["Description", "Tested an e-commerce web application similar to Amazon/Flipkart covering registration, login, product search, cart, checkout, and order confirmation workflows.", "", ""],
  ["Bullet 1", "Created and executed 55 manual test cases for registration, login, search, add to cart, checkout, and order confirmation modules.", "", ""],
  ["Bullet 2", "Identified and documented defects with severity, priority, reproduction steps, expected results, and actual results.", "", ""],
  ["Bullet 3", "Performed functional, regression, smoke, negative, and UI testing for critical e-commerce user journeys.", "", ""],
  ["Bullet 4", "Prepared test plan, bug report, mock JIRA tickets, RTM, and test execution summary.", "", ""],
  ["Tools", "Excel, JIRA / Mock JIRA, Browser DevTools", "", ""],
  ["Testing Types", "Functional, Regression, Smoke, Negative, UI, Boundary Value Testing", "", ""],
  ["Interview Summary", "Created a complete manual testing project for an e-commerce application and reported key defects across cart, checkout, payment, and order confirmation flows.", "", ""],
];
resume.getRange("A3:A11").format = { fill: colors.lightBlue, font: { bold: true } };
resume.getRange("B3:D11").merge(true);
resume.getRange("A3:D11").format.wrapText = true;
border(resume.getRange("A3:D11"));
widths(resume, [160, 300, 300, 300]);

const errors = await workbook.inspect({
  kind: "match",
  searchTerm: "#REF!|#DIV/0!|#VALUE!|#NAME\\?|#N/A",
  options: { useRegex: true, maxResults: 50 },
});
const previewDir = path.join(__dirname, "previews");
await fs.mkdir(previewDir, { recursive: true });
for (const sheetName of ["Dashboard", "Test Cases", "Bug Report", "Mock JIRA", "RTM", "Resume Points"]) {
  const preview = await workbook.render({ sheetName, autoCrop: "all", scale: 1, format: "png" });
  await fs.writeFile(path.join(previewDir, `${sheetName.replaceAll(" ", "_")}.png`), new Uint8Array(await preview.arrayBuffer()));
}
const output = await SpreadsheetFile.exportXlsx(workbook);
const xlsxPath = path.join(__dirname, "E-Commerce_Website_Testing_Project.xlsx");
await output.save(xlsxPath);
console.log(JSON.stringify({ xlsxPath, testCases: testCases.length, bugs: bugs.length, errors: errors.ndjson }, null, 2));
