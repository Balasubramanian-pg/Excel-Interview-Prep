### 185. BI: How do you calculate market basket analysis?

Market Basket Analysis is a data mining technique used in Business Intelligence (BI) to discover associations between products. It helps identify which items are frequently purchased together, forming the basis for strategies like product placement, cross-selling, and promotional bundling. The core metrics are Support, Confidence, and Lift.

> [!IMPORTANT]
> These calculations typically require a specific data structure. You need a list of unique transactions, and helper columns are often used to identify which items are present in each transaction. For example, you might have columns named `Has_Bread`, `Has_Milk`, etc., with TRUE/FALSE values for each transaction row.

#### 1. Support (Item Frequency)

Support measures the popularity of an item or a set of items. It is the proportion of total transactions in which the item appears.

**Formula:**
```excel
=COUNTIF(Transaction_Items, Item) / Total_Transactions
```

**How it works:**
*   `COUNTIF(Transaction_Items, Item)`: This counts the total number of times a specific `Item` appears in your dataset. `Transaction_Items` would be a range containing all items from all sales.
*   `Total_Transactions`: This is a cell containing the count of unique transactions.
*   The result is the percentage of all transactions that contain the specified item. A high support value means the item is very common.

#### 2. Confidence (A → B)

Confidence measures the likelihood of item B being purchased *given that* item A has been purchased. It is a conditional probability.

**Formula:**
```excel
=COUNTIFS(Trans_Has_A, TRUE, Trans_Has_B, TRUE) / COUNTIF(Trans_Has_A, TRUE)
```

**How it works:**
*   `COUNTIFS(Trans_Has_A, TRUE, Trans_Has_B, TRUE)`: This counts the number of transactions where **both** item A and item B are present. `Trans_Has_A` and `Trans_Has_B` are helper columns indicating the presence of each item for each transaction.
*   `COUNTIF(Trans_Has_A, TRUE)`: This counts the total number of transactions that contain item A.
*   The formula effectively says, "Of all the times A was bought, what percentage of the time was B also bought?"

> [!NOTE]
> Confidence can be misleading. If item B is extremely popular on its own (high support), the confidence for A→B might be high simply because B is in most baskets anyway, not because there is a strong relationship between A and B. This is where Lift becomes essential.

#### 3. Lift (A & B Association)

Lift measures how much more likely two items are to be purchased together than would be expected if they were statistically independent. It corrects for the individual popularity of each item.

**Formula:**
```excel
=Support_AB / (Support_A * Support_B)
```
Where `Support_AB` is the support for both items together, `Support_A` is the support for item A, and `Support_B` is the support for item B.

**Interpretation:**
Lift is the most important metric for determining the strength and nature of an association.

*   **`Lift > 1`**: A positive correlation exists. The items are purchased together more often than expected by chance. The presence of one item increases the likelihood of the other being purchased. (e.g., Chips and Salsa).
*   **`Lift = 1`**: There is no correlation. The items are statistically independent, and the purchase of one has no effect on the purchase of the other.
*   **`Lift < 1`**: A negative correlation exists. The items are purchased together less often than expected. The presence of one item may actually discourage the purchase of the other (e.g., competing brands).

> [!TIP]
> A Lift value of 2.5 means that customers are 2.5 times more likely to buy Item B when they buy Item A compared to the average customer. This is a strong indicator of a purchasing association.
