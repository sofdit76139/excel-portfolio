# My Excel Portfolio
Portfolio of various Excel workbooks and projects

## Most Excel workbooks contain explanations of the purpose of the workbook. 
No confidential data is included in any projects; all data has been altered.

##Other items:

1. Price Change Template.xlsb: Relies on pricelist_current.xlsx, which is a normalized template. The challenge in this situation is that all vendors will send you their data in whatever form, but for a script to work, you need it in a normalized format. This file accomplishes just that. The script writes to a change log file, also included (although the code will create it if it doesn't exist). The script will also produce files to be sent to the respective replenishment buyer to comment on, so items can be formally discontinued etc.
2. create_SKU_import_files.xlsb: After choosing a merchant to run the file for, the script will examine their working file, pull items from it, and created individual upload files by vendor. If certain data is not present, the script will inform the user and abort. It creates files for each replenishment buyer to let them know that is needed as an initial PO. It also creates a secondary file for pricing, which is uploaded once the SKU has been created in AX.
3. UPC_lookup.xlsx and NewSKUsYesterdayTodayYTD.xlsx: simple lookup tools that pull from the reporting data base.
