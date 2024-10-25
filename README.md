# Provisions and Write-Offs Management System

This project was developed to manage the entry, tracking, and reversal of provisions, making it easy to analyze which provisions have been partially or fully written off, while also enabling the registration of new clients and financial data. The entire database is stored and managed directly within Excel spreadsheets, providing convenience for data import and export.

**KEY FEATURES**

1. **Full CRUD for Provisions, Write-Offs, and Clients**  
   The system provides Create, Read, Update, and Delete (CRUD) functions to manage provisions, write-offs, and clients, enabling full control over financial data in the system, including the ability to edit and delete entries.

2. **Excel Data Import**  
   There's no need to enter provisions one by one manually. The system allows data import directly from a custom Excel template, making it easy to upload information in bulk.

3. **Provision and Write-Off Tracking**  
   The system offers detailed tracking of provisions, allowing users to check which have been fully written off, partially written off, or are still pending. The clear view enables easy control over completed and pending write-offs.

4. **Export of Pending Provisions**  
   Users can export a spreadsheet containing only pending provisions, along with details on amounts written off and those still outstanding. This helps in managing financial obligations and tracking amounts that need to be adjusted.

5. **Financial Data Handling and Formatting**  
   The system applies accounting formatting to all financial data, including gross revenue, taxes (ICMS, ISS, PIS, COFINS, CPRB), and written-off values, ensuring accuracy and clear organization for analysis.

6. **Reporting and Control**  
   Detailed reports can be generated to monitor provisions and write-offs, providing a clear and organized view of values and the status of each entry.

**TECHNOLOGIES USED**

1. **Flet**: Framework used for front-end development, providing a modern, user-friendly interface.
2. **Python**: Used for the system’s back-end, implementing business logic and Excel integration.
3. **Pandas & Openpyxl**: Essential libraries for data manipulation and interaction with Excel spreadsheets.
4. **Excel**: Used as the primary database to store provisions, write-offs, and clients, ensuring easy data integration and portability.

**CONCLUSION**

This project delivers a comprehensive management system for provisions and write-offs, with features that optimize financial control for companies. The ability to import and export data in Excel offers flexibility for those already using spreadsheets as a database. In addition, the complete CRUD functionality, detailed tracking of write-offs, and pending provisions export make this system a powerful tool for financial management.
