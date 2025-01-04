
# Discord Bot Documentation

## Table of Contents
- [Project Structure](#project-structure)
- [Link To Repo](#link-to-repo)
- [Purpose and Overview](#purpose-and-overview)
- [Features and Capabilities](#features-and-capabilities)
- [How It Works](#how-it-works)
- [Commands And Syntax](#commands-and-syntax)
- [Error Handling](#error-handling)
- [Data Integrity](#data-integrity)
- [Future Improvements](#future-improvements)
- [Conclusion](#conclusion)
  
---

## Project Structure
```bash
├── code
│   ├── main.py
│   └── keep_alive.py
├── data
│   ├── private
├── README.md

```

---

## Link to Repo: 
https://github.com/d-hoax/discord-bot-data-manager

---
---

## Purpose and Overview

### Purpose of the Bot
This bot is designed to simplify the management and interaction with data stored in an Excel spreadsheet. By leveraging Discord’s interactive commands, users can efficiently retrieve, update, and manage rows of data in the spreadsheet without needing direct access to it. The bot provides functionality for:

1. Searching for accounts based on specific criteria (e.g., rank).
2. Viewing details of specific accounts using unique identifiers (e.g., name).
3. Updating data for specific accounts or rows directly from Discord.
4. Ensuring real-time updates to the Excel sheet with minimal effort.

This bot serves as a bridge between Discord users and data management, allowing for seamless collaboration in a shared environment like gaming communities, team coordination, or organizational data tracking.

---

## Features and Capabilities

### Key Features
1. **Search Functionality**:
   - Quickly find accounts or data based on rank or other attributes.
   - Case-insensitive search for flexible queries.

2. **View Functionality**:
   - Retrieve detailed information about a specific account by name.
   - Display the rank or other relevant attributes of a given account.

3. **Update Functionality**:
   - Update specific cells in the spreadsheet using either the row number or account name.
   - Real-time data updates reflected directly in the Excel file.

4. **Data Integrity**:
   - Ensures valid column names and data ranges during updates.
   - Handles errors gracefully when invalid inputs are provided.

5. **User-Friendly Commands**:
   - Designed with straightforward syntax for ease of use.
   - Dynamic feedback and error messages to guide users in real-time.

---

## How It Works

### Underlying Functionality
The bot operates by:
1. **Loading the Spreadsheet**:
   - On startup, the bot loads the specified Excel file (`data.xlsx`) and checks for the worksheet (`Sheet1`).
   - If the file doesn’t exist, it creates a new Excel file with pre-defined column headers.

2. **Mapping Columns to Attributes**:
   - Each column in the spreadsheet corresponds to a specific attribute (e.g., `name`, `rank`, `username`).
   - A column mapping ensures that users can interact with the data using meaningful names instead of numerical indices.

3. **Handling Commands**:
   - The bot listens for user commands prefixed with `!` (e.g., `!search_rank`).
   - Based on the command and parameters provided, it processes the input and interacts with the spreadsheet accordingly.

4. **Saving Changes**:
   - After any update, the bot immediately saves the changes to the Excel file, ensuring data consistency.

---

## Commands and Syntax

### 1. **Search for Accounts by Rank**
**Command**: `!search_rank <rank>`

**Description**:
Searches the spreadsheet for all rows where the `rank` column matches the given rank exactly.

**Example Usage**:
```
!search_rank plat 1
```

**Expected Output**:
```
**Accounts with rank 'plat 1':**
Row 2: name=adi4386, tag=6338, rank=plat 1, username=adi4386, password=xyz, v?=n, email=n/a, sellable=n/a
Row 5: name=aufstrebend, tag=3200, rank=plat 1, username=aufstrebend, password=n/a, v?=n, email=n/a, sellable=n
```

### 2. **Show Rank of an Account by Name**
**Command**: `!show_rank <name>`

**Description**:
Retrieves and displays the rank of the account associated with the given `name`.

**Example Usage**:
```
!show_rank adi4386
```

**Expected Output**:
```
The rank for 'adi4386' is: **plat 1**
```

### 3. **Update a Cell by Row Number**
**Command**: `!update_cell <row_number> <column_name> <new_value>`

**Description**:
Updates the specified cell in the spreadsheet by directly referencing the row number and column name.

**Example Usage**:
```
!update_cell 2 rank ascendant 3
```

**Expected Output**:
```
Updated row 2, column 'rank' to 'ascendant 3'.
```

### 4. **Update a Cell by Name**
**Command**: `!update_name <name> <column_name> <new_value>`

**Description**:
Updates a specific cell in the row where the `name` column matches the given name.

**Example Usage**:
```
!update_name adi4386 rank ascendant 3
```

**Expected Output**:
```
Updated row 2 where name='adi4386', column 'rank' to 'ascendant 3'.
```

---

## Error Handling

1. **Invalid Column Name**:
   - If the user provides a column name that doesn’t exist, the bot will return:
    ```
    Invalid column 'invalid_column'. Valid columns: name, tag, rank, username, password, v?, email, sellable
    ```

2. **Invalid Row Number**:
   - If the specified row number is out of range, the bot will return:
    ```
    Row number 999 is out of range (2 - 50).
    ```

3. **No Match Found**:
   - If no row matches the given name or rank, the bot will return a message such as:
    ```
    No rows found with rank exactly matching 'invalid_rank'.
    ```
     or
    ```
    No row found with name exactly matching 'invalid_name'.
    ```

---

## Data Integrity

The bot ensures:
1. **Case-Insensitive Matching**:
   - Commands like `!show_rank` and `!update_name` handle case-insensitive comparisons for better user experience.

2. **Validation of Column Names**:
   - Updates are only allowed for valid columns as defined in the `COLUMN_MAP`.

3. **Immediate Saving**:
   - Every update is immediately saved to the Excel file to prevent data loss.

4. **Error Feedback**:
   - Users are notified of any invalid inputs with clear error messages.

---

## Future Improvements

1. **Support for Partial Matches**:
   - Extend search functionality to include partial matches for names or ranks.

2. **Duplicate Handling**:
   - Add functionality to handle duplicate names or allow users to specify additional criteria (e.g., `name` + `tag`).

3. **Enhanced Security**:
   - Mask sensitive data like passwords in the output.

4. **Rich Output**:
   - Use Discord embeds for prettier and more structured responses.

5. **Additional Commands**:
   - Add commands for deleting rows, listing all data, or exporting the spreadsheet.

---

## Conclusion

This bot provides a simple yet powerful interface for managing data stored in Excel spreadsheets via Discord. It bridges the gap between technical and non-technical users, allowing real-time collaboration and updates in a familiar chat environment. With features like search, view, and update, the bot is versatile enough to cater to various use cases, including gaming communities, project management, and team coordination. Its design ensures ease of use, data integrity, and flexibility for future enhancements.
