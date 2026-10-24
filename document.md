# Mini ERP System in Excel

## 1. Project Overview

The Mini ERP System in Excel is an Excel-based business management solution developed to simulate core Enterprise Resource Planning (ERP) workflows in a small business environment.

The system integrates business document management, inventory control, customer and vendor management, financial transactions, and operational reporting within a single Excel application.

It uses VBA automation, UserForms, structured Excel tables, formulas, PivotTables, and charts to support business operations and maintain consistency across related processes.

The project focuses on demonstrating how Excel can be used to build an integrated business system with document traceability, inventory updates, financial balance management, and validation controls.

## 2. Project Objectives

The main objectives of the project are to:

- Develop an integrated business workflow for sales, purchasing, inventory, and financial operations.
- Maintain traceability between related business documents.
- Track outstanding balances and financial transactions.
- Manage inventory changes resulting from relevant business operations.
- Apply validation rules to prevent invalid operations.
- Provide a centralized interface for managing business records.
- Present summarized business information through a dashboard.

## 3. Project Scope

The system covers the following functional areas:

- **Business document management:** Creation and processing of quotations, sales orders, delivery notes, invoices, purchase orders, goods receipts, vendor bills, vendor credits, and credit notes.
- **Entity management:** Maintenance of customer, vendor, and employee records.
- **Item and inventory management:** Maintenance of item information, prices, stock quantities, stock thresholds, and stock-related status indicators.
- **Financial operations:** Management of customer collections, vendor payments, credit applications, and related reversals.
- **Document traceability:** Relationships between source documents and documents generated from them.
- **Validation and status management:** Controls for supported operations and indicators reflecting document and inventory states.
- **Business reporting:** Dashboard charts based on PivotTables.

The project is intended as a functional ERP-style simulation built in Excel, rather than a replacement for a full-scale commercial ERP platform.

## 4. Technology Stack

| Technology | Purpose |
|---|---|
| Microsoft Excel | Application environment and data storage |
| VBA and macros | Business logic, automation, and workflow processing |
| UserForms | Structured forms for data entry and transaction processing |
| Excel Tables | Storage and organization of business records |
| Excel formulas | Calculations and derived information |
| PivotTables and charts | Summarized reporting and dashboard visualization |
| Shapes and interface elements | Navigation, actions, and visual status indicators |

## 5. System Architecture Overview

The application is organized into six main modules: ADMIN, ORDERS, PAYMENTS, NAMES, ITEMS, and DASHBOARD.

Each module serves a specific operational purpose while contributing to the overall business workflow.

The system's central data structures support document management, document line items, financial transactions, payment and credit applications, and inventory information.

### 5.1 ADMIN — System Configuration

The ADMIN module centralizes configurable parameters used throughout the system.

Its configuration includes:

- **Business information:** Company name and address details.
- **Document messages:** Configurable text associated with each document type.
- **Tax configuration:** Tax rate information used in business operations.
- **Document numbering:** Document prefixes, such as QUO, SO, DN, and INV, and reference numbering settings.
- **Document definitions:** Descriptions associated with each document type.
- **Item and account categories:** Available item types, expense types, and sales types.

This module allows key business settings to be maintained centrally.

### 5.2 ORDERS — Business Document Management

The ORDERS module manages the creation and processing of business documents, from quotations to credit notes.

Users can select the relevant entity, add items and quantities, and perform document-specific operations through an action panel.

The action panel groups its functionality into four categories.

**General Actions**

- Save
- Add New
- Cancel
- Print
- Search Order

**Next Action**

Provides the next available document-processing step based on the loaded document and its supported workflow.

For example, loading a Sales Order can provide the option to create a Delivery Note, allowing the user to continue the document lifecycle.

**Create**

Provides options to create related documents, including:

- Credit Note
- Vendor Credit
- Goods Receipt
- Vendor Bill

**Apply**

Allows previously created credits to be applied to eligible business documents through the available actions:

- Credit Note
- Vendor Bill

The module also displays informational badges that help users understand the current state of a document, including:

- **Status:** A visual indicator of the document's operational or financial state.
- **Balance:** The document's relevant outstanding or available balance.
- **Source:** Information about the source document and its traceability.
- **Credit settlement indicators:** Visual information indicating whether an invoice or vendor bill has been settled through credit applications.

Status badges do not always display the raw status stored in the Orders table. For example, an invoice's visual status is determined from its balance, while a Sales Order's badge reflects its stored operational status.

These features bring document processing, operational status, and document lineage together in a single interface.

### 5.3 PAYMENTS — Financial Transaction Management

The PAYMENTS module manages collections from customers and payments to vendors associated with invoices and vendor bills.

Its interface includes:

- A transaction history table displaying previously recorded financial transactions.
- A dedicated UserForm for processing collections and payments.
- Search and filtering controls for locating transaction records.

The module validates payment amounts against the selected document's outstanding balance and creates a corresponding application record. It also supports payment reversals, which update the related application records and trigger balance recalculation.

These functions provide visibility into recorded financial activity and help users monitor transaction history and outstanding document balances.

### 5.4 NAMES — Entity Management

The NAMES module manages the business entities maintained by the system, including customers, vendors, and employees.

Its interface follows a structure similar to PAYMENTS, with:

- A records table.
- A dedicated UserForm for entity management.
- Search and filtering controls.

This module centralizes entity information used throughout business document and transaction workflows.

### 5.5 ITEMS — Product and Inventory Management

The ITEMS module manages the products and other items available for use in business operations.

Its interface includes a records table, a dedicated UserForm, and search and filtering controls.

The records table displays item-specific information, including:

- Item cost and sales price.
- Current stock quantity.
- Minimum and maximum stock thresholds.
- Item status and stock-related status indicators.

The current stock quantity, represented by `Qty On Hand`, is calculated from the inventory movement records stored in `OrderItems`, rather than being directly edited through the item maintenance form.

The `Stock Status` field is derived from the item's type, current stock quantity, and configured minimum and maximum stock thresholds. This provides a visual classification of inventory conditions, including critical stock levels, reorder conditions, warning levels, healthy stock, and surplus stock.

Items can be maintained and filtered through the module's interface. Deletion rules also preserve inventory and historical integrity: items with nonzero stock cannot be deleted, and items with transaction history are retained as inactive records rather than physically deleted.

This module provides inventory visibility and supports operational workflows in which stock changes through supported document processes.

### 5.6 DASHBOARD — Business Overview

The DASHBOARD module provides a summarized view of business activity through charts built from PivotTables.

It is intended to help users quickly interpret operational information without reviewing individual records across the system.

The interface includes:

- A data refresh button to update the displayed information.
- A navigation button to return to the ERP's main interface.

The dashboard complements the operational modules by presenting business information in a more accessible visual format.

## 6. Core Data Model

The system uses five central Excel tables to organize business documents, document line items, financial transactions, applications, and inventory information.

These tables support the system's main operational workflows and provide the data required to track document relationships, financial activity, and inventory changes.

### 6.1 Main Tables

#### Orders

The `Orders` table stores business documents, including their identifiers, reference numbers, types, associated entities, dates, financial amounts, statuses, and relationships to other documents.

Key fields include:

- `Order ID`: Document identifier.
- `Ref.`: Document reference number.
- `Order Type`: Document classification.
- `Name`, `Address 1`, and `Address 2`: Entity and address information associated with the document.
- `Inc./Exp. Type`: Income or expense classification.
- `Date` and `Due Date`: Document and due dates.
- `Subtotal`, `TaxAmount`, and `TotalFinal`: Financial amount fields.
- `Payments`, `Adjustments`, and `Total Applied`: Fields associated with payments, credits, and total applied amounts.
- `Balance`: Outstanding balance or remaining available credit, depending on the document type.
- `Status`: Stored document status.
- `ParentOrder ID` and `SourceSOID`: Fields supporting document relationships and traceability to the originating Sales Order.

This table is central to document management and provides the information needed to track document states, financial amounts, and relationships between business documents.

#### OrderItems

The `OrderItems` table stores the item lines associated with business documents.

Key fields include:

- `Order ID`: Identifier of the associated document.
- `Item ID` and `Item Name`: Item identification.
- `Item Desc.`: Item description recorded on the document line.
- `Qty`: Quantity associated with the line.
- `UnitPrice` and `LineTotal`: Unit price and line amount.
- `Movement Qty`: Quantity field associated with physical inventory movements.
- `Item Type`: Item classification.
- `Date`: Date recorded for the line.
- `SourceOrderID` and `SourceItemID`: Fields supporting source references between related document lines.

This table provides item-level detail for business documents and supports the tracking of quantities and source relationships across relevant workflows.

The `Movement Qty` field is also used to calculate current inventory. Its values can represent stock decreases, stock increases, or zero movement, depending on the document type and implemented business rules. Compensating movement lines are also recorded when the reviewed inventory reversal routine is used.

#### Transactions

The `Transactions` table stores financial transaction records associated with customer collections and vendor payments.

Key fields include:

- `Trans. ID`: Transaction identifier.
- `Trans. Type`: Transaction classification.
- `Name`: Associated entity.
- `Date`: Transaction date.
- `Account`: Account information.
- `Document ID`: Identifier of the associated document.
- `Reference`: Transaction reference.
- `Amount`: Transaction amount.
- `Description`: Transaction details.
- `Status`: Transaction status, including supported reversal states.

These records support financial history and provide a way to identify transactions associated with business documents.

#### Applications

The `Applications` table records financial applications between source and destination records. These include payment applications and credit applications.

Key fields include:

- `AppID`: Application identifier.
- `FromID`, `FromDocNo`, and `FromType`: Identification of the source transaction or document.
- `ToID`, `ToDocNo`, and `ToType`: Identification of the destination document.
- `Amount`: Amount applied.
- `Status`: Application status, such as Active or Reversed.
- `AppliedDate`: Application date.
- `CreatedBy`: User or process recorded as creating the application.
- `Currency`: Currency information.
- `Memo`: Additional notes.

The table supports the tracking of applied amounts and their relationships to source and destination records. Active applications contribute to financial balance calculations, while reversed applications are excluded from those calculations.

#### Items

The `Items` table stores item information and inventory-related data.

Key fields include:

- `Company`: Company information associated with the item record.
- `Primary Vendor`: Primary vendor information.
- `Item ID` and `Item Name`: Item identification.
- `Item Type`: Item classification.
- `Cost/Amount` and `Sales Price`: Item cost and sales price information.
- `Qty On Hand`: Current stock quantity calculated from inventory movement records.
- `Min. Stock` and `Max. Stock`: Stock threshold fields.
- `Status`: Item status, such as Active or Inactive.
- `Stock Status`: Derived inventory classification based on item type, current stock, and configured thresholds.

The table supports inventory visibility and provides item information used throughout relevant business workflows.

### 6.2 Data Relationships

The system's data model connects business documents with their item lines, financial transactions, applications, and inventory records.

#### Document Relationships

The `Orders` table contains `ParentOrder ID` and `SourceSOID`, which support relationships between documents and traceability to the originating Sales Order.

These fields help preserve document lineage across supported business workflows.

#### Document Line Relationships

The `OrderItems` table associates item lines with business documents through `Order ID`.

The fields `SourceOrderID` and `SourceItemID` provide additional source references for supported relationships between document lines.

Together, these fields support the association of document lines with their corresponding documents and the tracking of item-level relationships across relevant workflows.

#### Financial Transaction Relationships

The `Transactions` table includes `Document ID`, which identifies the business document associated with a financial transaction.

Payment processing also creates a corresponding record in `Applications`. This allows the system to associate the payment with the target invoice or vendor bill and use active applications in the balance calculation.

#### Application Relationships

The `Applications` table identifies source and target records through `FromID` and `ToID`, together with their corresponding document numbers and types.

For payment applications, the source represents the payment transaction and the destination represents the invoice or vendor bill. For credit applications, the source represents the credit document and the destination represents the document receiving the credit.

The `Amount` and `Status` fields support the tracking of application amounts and their current states, including supported reversals.

#### Inventory Relationships

The `Items` table stores item information and current stock quantities, while `OrderItems` records item quantities and includes the `Movement Qty` field.

The `Qty On Hand` formula sums the `Movement Qty` values associated with each `Item ID`:

```excel
=SUMIFS(OrderItems[Movement Qty];OrderItems[Item ID];$C4)
```

Consequently, the calculated stock quantity depends on the inventory movements recorded in `OrderItems`. A negative movement reduces stock, a positive movement increases stock, and a zero movement does not change the calculated quantity.

The `Stock Status` field uses the calculated quantity and the item's minimum and maximum stock thresholds to classify its inventory condition.

#### Overall Data Model

Together, the five tables support the system's main business processes:

- Document creation and traceability.
- Item-level document management.
- Financial transaction recording.
- Payment and credit application tracking.
- Inventory monitoring and stock-related operations.

The relationships described above are based on the available table structure and reviewed VBA logic. Implementation-specific rules, including balance calculations, stock movement conditions, status transitions, and reversal behavior, are addressed in the relevant business process sections.

## 7. Document Status, Traceability, and Financial Control

The ORDERS module manages the lifecycle of business documents and maintains relationships between source documents and documents derived from them. The system also tracks document statuses, outstanding balances, payments, and credit applications to support operational and financial control.

This section describes the main sales workflow, the rules used to calculate balances and display document statuses, and the behavior of supported payment and credit reversals.

### 7.1 Document Lifecycle and Traceability

The main sales workflow connects the following stages:

**Sales Order → Delivery Note → Invoice → Payment**

Each document serves a different operational purpose:

- **Sales Order (SO):** Records the customer's order and serves as the originating document for the subsequent sales workflow.
- **Delivery Note (DN):** Records the delivery stage and supports the corresponding inventory movement.
- **Invoice (INV):** Records the amount owed by the customer and supports accounts receivable tracking.
- **Payment:** Records a collection against an eligible invoice and contributes to reducing its outstanding balance.

The system maintains document relationships through fields such as `ParentOrder ID` and `SourceSOID` in the `Orders` table. These fields support the identification of related documents and the traceability of documents back to their originating Sales Order.

The Sales Order's stored operational status also helps determine whether the next processing step is available. For example, the interface allows the creation of a Delivery Note when the Sales Order is Active or Partial. A Delivery Note with an eligible status can then proceed to invoice creation.

The end-to-end sales workflow, including document creation, inventory changes, payment processing, balance reconciliation, and traceability, was tested as part of the project's validation process.

### 7.2 Document Status and Outstanding Balances

The system maintains stored document statuses and financial balances. The interface may translate those values into visual badges that communicate a document's operational or financial state.

#### Stored Document Statuses

The reviewed status-management logic uses operational states such as Active, Partial, Closed, and Cancelled for supported sales and purchasing documents. Credit Note and Vendor Credit records use states such as Unapplied, Partial, Applied, and Cancelled.

#### Visual Status Indicators

The `RefreshDocumentStatus` procedure derives several badges from the document type and its data:

- **Sales Order:** Displays Active, Partial, Completed, or Cancelled according to its stored operational status.
- **Delivery Note:** Displays Posted.
- **Invoice:** Displays Paid when the balance equals zero, Partial when the balance is below the total but not zero, and Unpaid when the balance equals the total.
- **Credit Note and Vendor Credit:** Display Applied when the balance equals zero, Partial when some credit has been used, and Unapplied when the balance equals the total.
- **Quote:** Displays Converted when an associated Sales Order exists, or Draft otherwise.

These badges are visual representations of the implemented rules and should not be interpreted as a direct copy of the stored status for every document type.

#### Balance Recalculation

The `RecalcularSaldos` procedure rebuilds the financial amounts for invoices, vendor bills, credit notes, and vendor credits using the `Applications` table.

Only applications with an Active status are included in the calculation. For invoices and vendor bills, the destination document identifier (`ToID`) determines which applications contribute to the document. For credit notes and vendor credits, the source document identifier (`FromID`) determines which applications contribute to the credit document.

The calculation follows these rules:

- **Payments:** The sum of active applications whose source type is Payment.
- **Adjustments:** The sum of active applications whose source type is not Payment, representing credit applications in the reviewed logic.
- **Total Applied:** The total of active applications associated with the document.
- **Balance:** `TotalFinal - Total Applied`.

Consequently, reversing an application removes its amount from the active application total. Recalculating balances then restores the corresponding outstanding invoice balance or available credit, as applicable.

#### Payment Validation

The PAYMENTS form checks that the payment amount is numeric, greater than zero, and does not exceed the selected document's outstanding balance, subject to the implemented tolerance. The form also restricts payment processing to invoices and vendor bills.

The rejection of a payment amount exceeding the selected document's outstanding balance was verified during testing.

### 7.3 Credit Notes and Credit Applications

The system supports Credit Notes and Vendor Credits, as well as applying available credits to eligible invoices and vendor bills.

#### Credit Note Creation and Inventory

A Credit Note records a credit associated with a customer transaction. Depending on the selected restocking option and whether the original workflow generated a physical inventory movement, creating a Credit Note can also restore stock.

The reviewed logic generates a positive inventory movement for eligible product lines when restocking is enabled and the original workflow moved stock. Otherwise, the Credit Note does not generate a stock movement through this routine.

The tested workflow used a Credit Note with a physical inventory movement. The inventory state was checked during the credit process and its subsequent reversal.

#### Credit Application Records

The `Applications` table records the source document, destination document, amount, status, date, and related metadata for each application.

When a credit is applied, its application record is created with an Active status. The balance recalculation procedure then includes that application when calculating the remaining balance of the affected invoice and the available balance of the credit document.

The visual status of a credit document is derived from its remaining balance:

- **Unapplied:** The full credit amount remains available.
- **Partial:** Some of the credit has been applied, but a balance remains.
- **Applied:** No credit balance remains available.

The interface also provides contextual actions for applying an available credit or viewing its applications when the credit has been fully used.

#### Credit Application Validation

The system validates credit availability before applying a credit. Testing confirmed that an application exceeding the available credit is rejected and that a valid application updates the corresponding financial balances and visual statuses.

### 7.4 Reversals and Financial Control

The system supports the reversal of recorded payments and credit applications. Reversals change the status of the relevant records rather than simply deleting their history.

#### Payment Reversal

When a payment is reversed through the PAYMENTS form, the associated transaction is marked Reversed and its active applications are reversed. The system then recalculates the related document balances and updates the document status.

Because reversed applications are excluded from the balance calculation, the affected invoice or vendor bill regains the outstanding balance associated with the reversed payment.

Testing confirmed that reversing a payment restores the invoice's outstanding balance.

#### Credit Application Reversal

The reviewed code includes a procedure for reversing a specific application. It marks the application as Reversed, invokes a physical movement reversal routine, recalculates financial balances, and updates the source document's stored status.

The physical movement routine adds compensating lines to `OrderItems` for the movement quantities associated with the source credit document. These lines use the opposite movement quantity, allowing the inventory calculation to account for the reversal without deleting the original movement records.

This behavior is distinct from the general routine that reverses all active applications associated with a source document, which changes application statuses but does not itself invoke the physical inventory reversal routine.

In the demonstrated credit application reversal workflow, the source Credit Note had an inventory movement. The reversal affected stock through the compensating movement logic, and the final inventory state was checked as part of the test sequence.

#### Financial and Operational Consistency

The reversal workflow demonstrates how the system can preserve transaction history while recalculating financial balances and, in the tested credit scenario, compensating for a related inventory movement.

The documented behavior reflects the specific reversal routines and test scenarios reviewed. It should not be interpreted as meaning that every financial application reversal automatically changes inventory in every workflow.

### 7.5 Validation Evidence

The following test cases provide evidence for the behaviors described in this section:

| Test Case | Verified Behavior |
|---|---|
| TC-01 | Sales Order, Delivery Note, and Invoice traceability |
| TC-02 | Payment processing through the PAYMENTS form |
| TC-05 | Payment reversal and restoration of the outstanding balance |
| TC-07 | Credit application validation and successful application |
| TC-08 | Credit application reversal, resulting document statuses, and final inventory verification |
| TC-09 | Rejection of a payment amount exceeding the outstanding balance |

See [Section 10 — Testing and Validation](#10-testing-and-validation) for the overall test results and [Section 11 — Project Evidence](#11-project-evidence) for the supporting demonstrations.

## 8. Inventory Management

The system maintains item information in the `Items` table and uses document line records in `OrderItems` to calculate stock quantities and support inventory-related operations.

### 8.1 Stock Calculation

The `Qty On Hand` field is calculated using an Excel formula that sums the `Movement Qty` values associated with each item:

```excel
=SUMIFS(OrderItems[Movement Qty];OrderItems[Item ID];$C4)
```

The resulting quantity reflects the net inventory movements recorded for the corresponding `Item ID`.

This approach allows the system to derive stock quantities from recorded movements rather than relying on direct manual changes to the current stock field.

### 8.2 Stock Status Classification

The `Stock Status` field is calculated from the item type, current stock quantity, minimum stock threshold, and maximum stock threshold.

For product items, the implemented formula assigns the following classifications:

| Condition | Stock Status | Meaning |
|---|---|---|
| Current stock equals zero | Stockout / Critical | No stock is available |
| Current stock is less than or equal to minimum stock | Low Stock / Reorder | Replenishment is indicated |
| Current stock is above minimum and within the next 35% of the minimum-to-maximum range | Warning / Buffering | Stock is approaching the healthy range |
| Current stock is above the warning threshold and less than or equal to maximum stock | Healthy / In Stock | Stock is within the healthy range |
| Current stock exceeds maximum stock | Overstock / Surplus | Stock exceeds the configured maximum |

Items classified as services display `-` instead of an inventory status.

These classifications are formula-derived indicators based on the configured thresholds. They provide an at-a-glance view of inventory conditions rather than representing separate inventory transactions.

### 8.3 Inventory Movements by Document Type

Inventory movements depend on the document type and the implemented processing rules.

#### Delivery Note

During the reviewed sales processing workflow, a Delivery Note records a negative `Movement Qty` for the products delivered. The process checks stock availability before saving the document. The resulting movement reduces the calculated `Qty On Hand`.

#### Invoice

The reviewed invoice-processing routine records `Movement Qty = 0`. The invoice records the billed quantities and financial amounts without generating a second stock movement in that workflow.

#### Customer Credit Note

The Credit Note form provides a `Restock Inventory` option, enabled by default when the form is initialized.

When restocking is enabled, the routine checks whether the original sales workflow moved stock. If the condition is satisfied, eligible product lines receive a positive `Movement Qty`, increasing the calculated stock quantity. If the option is disabled or the original workflow did not move stock, the routine records no physical stock movement for those lines.

#### Vendor Credit

The Vendor Credit form provides a `Return to Vendor` option, also enabled by default.

When the option is enabled, eligible product lines receive a negative `Movement Qty` through the reviewed routine, representing stock returned to the vendor. When the option is disabled, the routine records no physical stock movement for those lines.

These rules distinguish financial credit adjustments from physical inventory movements. A credit document does not necessarily change stock; the result depends on the document type, selected option, item type, and implemented conditions.

### 8.4 Inventory Validation and Item Integrity

The item management module maintains item descriptions, types, vendor information, cost and sales prices, and minimum and maximum stock thresholds.

The item maintenance form does not directly edit the formula-driven `Qty On Hand` field.

The system also applies deletion restrictions to help preserve inventory and historical integrity:

- An item with nonzero stock cannot be deleted.
- An item with existing document history can be marked Inactive instead of being physically deleted.
- An item with no document history and zero stock can be physically deleted.

These rules help retain historical references while preventing the removal of items that still have recorded stock or transaction history.

### 8.5 Inventory Reversal Behavior

The reviewed credit application reversal routine can compensate for an inventory movement associated with the source credit document.

It adds new `OrderItems` lines with movement quantities opposite to the original movements, allowing the formula-driven stock calculation to account for the reversal without deleting the original records.

This behavior applies to the reviewed reversal routine and scenario. It should not be generalized to every payment reversal, credit application reversal, or financial adjustment in the system.

### 8.6 Inventory Validation Evidence

The following test cases support the inventory behavior described in this section:

- **TC-03:** Inventory reduction after Delivery Note processing.
- **TC-06:** Credit Note creation and inventory restoration.
- **TC-08:** Credit application reversal and final inventory verification in the tested scenario.

The demonstration videos provide additional visual evidence of the relevant stock changes and the associated document workflows.

## 9. Dashboard and Reporting

The DASHBOARD module provides a summarized view of business activity through charts built from PivotTables.

The dashboard is intended to provide a high-level overview of business information without requiring users to inspect individual records across the operational modules.

Its interface includes a refresh button to update the displayed information and a navigation button to return to the main ERP interface.

During testing, the dashboard was reviewed before and after selected business operations to verify that the displayed information reflected the tested changes after the relevant refresh process.

The dashboard complements the detailed records maintained in the system by presenting selected business information in a visual format.

## 10. Testing and Validation

The system's functional behavior is evaluated through documented test cases covering business workflows, financial operations, inventory changes, validation rules, and document traceability.

The test suite includes the following cases:

| Test Case | Test Focus | Result |
|---|---|---|
| TC-01 | Sales Order → Delivery Note → Invoice traceability | Pass |
| TC-02 | Payment processing through the PAYMENTS form | Pass |
| TC-03 | Inventory reduction after Delivery Note processing | Pass |
| TC-04 | Dashboard information before and after operations | Pass |
| TC-05 | Payment reversal and restoration of the outstanding balance | Pass |
| TC-06 | Credit Note creation and inventory restoration | Pass |
| TC-07 | Credit application validation and successful application | Pass |
| TC-08 | Credit application reversal, resulting document statuses, and final inventory verification | Pass |
| TC-09 | Rejection of a payment amount exceeding the outstanding balance | Pass |

**Test results: 9 of 9 test cases passed.**

The test results reflect the observed outcomes of the executed test scenarios. The project's test case matrix records the expected result, observed result, execution status, and supporting evidence. The public video demonstrations provide additional visual evidence of the tested workflows.

## 11. Project Evidence

### Video Demonstrations

The following unlisted YouTube videos demonstrate the main operational workflows, financial controls, inventory updates, and document traceability of the system.

#### 1. Main Sales Workflow — SO → DN → INV → Payment

Demonstrates the end-to-end sales process, including Sales Order processing, Delivery Note creation, inventory reduction, Invoice generation, payment application, balance reconciliation, dashboard updates, and document traceability.

[Watch the Main Sales Workflow](https://youtu.be/c0g0aNYOZcw)

#### 2. Payment Reversal & Credit Note Workflow

Demonstrates payment reversal, restoration of the outstanding invoice balance, Credit Note creation and inventory restoration, credit application validation and settlement, application reversal, and final state verification.

[Watch the Payment Reversal & Credit Note Workflow](https://youtu.be/ZTDzejFjyL0)

*Both videos are unlisted on YouTube and can be accessed by anyone with the links.*

### Test Documentation

The project includes a test case matrix covering nine functional test cases (TC-01 to TC-09). Each test records the expected result, observed result, execution status, and supporting evidence.

All nine test cases passed. The video demonstrations provide additional visual evidence of the tested workflows and their resulting document, financial, and inventory states.

## 12. Scope and Limitations

The system is an educational and portfolio-oriented ERP simulation implemented in Microsoft Excel.

Its capabilities and limitations are determined by the workbook's implemented business rules, VBA automation, Excel data structures, and supported workflows.

It should not be considered a full-scale commercial ERP or a substitute for a production system with dedicated database infrastructure, multi-user access controls, and enterprise-level security.

The public repository initially focuses on documentation and demonstration evidence. The Excel workbook itself is not included in the initial public release.

## 13. Conclusion

The Mini ERP System in Excel demonstrates how Excel, VBA, and structured data tables can be combined to simulate integrated business operations.

By connecting document management, inventory control, financial transactions, payment and credit applications, and reporting, the project illustrates the importance of data relationships, operational traceability, validation, and consistent business state management.

The project also demonstrates how transaction reversals can be handled through status changes, balance recalculation, and compensating inventory movements in supported workflows, providing a practical example of business process automation within an Excel-based ERP simulation.
