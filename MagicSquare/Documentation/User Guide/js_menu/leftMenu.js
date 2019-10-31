//Each menuItem in the menu is a menuItem object whose parameters represent the 
//menuItem's text and link properties. 
//The code positions the menu on the screen and then renders it in the appropriate 
//place.

function displayLeftMenu(folderParentImages, folderParentFile)
{
	  initLeftMenu(folderParentImages); 
	  var mnuIntroduction = new menu('mnuIntroduction', 
         new menuTrigger('trgIntroduction','Introduction'),
         new menuItem('Introduction', folderParentFile + '../../basics/about-easyxls-excel-library.html'));  
      mnuIntroduction.position(-1,10,230);
      mnuIntroduction.write(); 
      
      var mnuLicensing = new menu('mnuLicensing', 
         new menuTrigger('trgLicensing','Licensing'),
         new menuItem('Licensing', folderParentFile + '../../licensing/easyxls-license-agreement.html'), 
		 new menuItem('License types', folderParentFile + '../../licensing/easyxls-licenses.html')) ;

      mnuLicensing.position(-1,10,230);
      mnuLicensing.write();   
      
      var mnuGettingStarted = new menu('mnuGettingStarted', 
         new menuTrigger('trgGettingStarted','Getting started'),
         new menuItem('EasyXLS - .NET Excel Library', folderParentFile + '../../getting-started/easyxls-dot-net-excel-library.html'), 
         new menuItem('EasyXLS - COM+ Excel Component', folderParentFile + '../../getting-started/easyxls-com-excel-component.html'),
         new menuItem('EasyXLS - Java Excel Library', folderParentFile + '../../getting-started/easyxls-java-excel-library.html')) ; 
      mnuGettingStarted.position(-1,10,230);
      mnuGettingStarted.write(); 
      
      var mnuImportFiles = new menu('mnuImportFiles', 
         new menuTrigger('trgLoadFiles','Import files'),
         //new menuItem('Import from Datasets/Resultsets',folderParentFile + 'basics/importFromDataSetsResultSets.html'),
         new menuItem('Import from XLS file',folderParentFile + '../../basics/import-from-xls-file-format.html'),
       	 new menuItem('Import from XLSX file',folderParentFile + '../../basics/import-from-xlsx-file-format.html'),
       	 new menuItem('Import from XLSM file',folderParentFile + '../../basics/import-from-xlsm-file-format.html'),
       	 new menuItem('Import from XLSB file',folderParentFile + '../../basics/import-from-xlsb-file-format.html'),
	     new menuItem('Import from XML Spreadsheet file', folderParentFile + '../../basics/import-from-xml-spreadsheet-file-format.html'),
	     new menuItem('Import XML data', folderParentFile + '../../basics/import-data-in-xml-format.html'),
         new menuItem('Import from TXT file', folderParentFile + '../../basics/import-from-txt-file-format.html'),
         new menuItem('Import from CSV file', folderParentFile + '../../basics/import-from-csv-file-format.html'),
         new menuItem('Import from HTML file', folderParentFile + '../../basics/import-from-html-file-format.html'),  
         new menuItem('Import Excel file to DataSet',folderParentFile + '../../basics/import-excel-to-dataset.html'),   
         new menuItem('Import Excel file to ResultSet',folderParentFile + '../../basics/import-excel-to-resultset.html'),
         new menuItem('Import Excel file to List',folderParentFile + '../../basics/import-excel-to-list.html')
         //new menuItem('Import Excel file in datasets/resultsets',folderParentFile + 'basics/exportDataSetsResultSets.html'),
         );            

      var mnuExportFiles = new menu('mnuExportFiles', 
         new menuTrigger('trgExportFiles','Export files'),
         new menuItem('Export XLS file',folderParentFile + '../../basics/export-to-xls-file-format.html'),
		 new menuItem('Export XLSX file',folderParentFile + '../../basics/export-to-xlsx-file-format.html'),
		 new menuItem('Export XLSM file',folderParentFile + '../../basics/export-to-xlsm-file-format.html'),
		 new menuItem('Export XLSB file',folderParentFile + '../../basics/export-to-xlsb-file-format.html'),
         new menuItem('Export XML Spreadsheet file', folderParentFile + '../../basics/export-to-xml-spreadsheet-file-format.html'),
         new menuItem('Export TXT file', folderParentFile + '../../basics/export-to-txt-file-format.html'),
         new menuItem('Export CSV file', folderParentFile + '../../basics/export-to-csv-file-format.html'),
         new menuItem('Export HTML file', folderParentFile + '../../basics/export-to-html-file-format.html'),
         new menuItem('Export DataSet to Excel file',folderParentFile + '../../basics/export-dataset-to-excel.html'),
         new menuItem('Export ResultSet to Excel file',folderParentFile + '../../basics/export-resultset-to-excel.html'),
         new menuItem('Export List to Excel file',folderParentFile + '../../basics/export-list-to-excel.html')
         ); 
         
      var mnuSheets = new menu('mnuSheets', 
         new menuTrigger('trgSheets','Sheets'),
         new menuItem('Create sheets',folderParentFile + '../../basics/create-excel-file-multiple-sheets.html'),
         new menuItem('Sheet properties',folderParentFile + '../../basics/sheet-properties.html'));
         
      var mnuAddressingCells = new menu('mnuAddressingCells', 
         new menuTrigger('trgAddressingCells','Addressing cells'),
         new menuItem('Cell values',folderParentFile + '../../basics/import-export-excel-data.html'),
         new menuItem('Formulas',folderParentFile + '../../basics/import-export-excel-formulas.html'),
         new menuItem('Calculate formulas',folderParentFile + '../../basics/excel-calculation-engine.html'));
         
      var mnuAutoformatting = new menu('mnuAutoformatting', 
         new menuTrigger('trgAutoformatting','Autoformatting'),
         new menuItem('Predefined formatting for cell ranges',folderParentFile + '../../basics/excel-predefined-formatting.html'),       
         new menuItem('Custom formatting for cell ranges',folderParentFile + '../../basics/excel-custom-formatting.html')); 
         
      var mnuCellFormatting = new menu('mnuCellFormatting', 
         new menuTrigger('trgCellFormatting','Cell formatting'),
         new menuItem('Formatting cells, columns and rows',folderParentFile + '../../basics/format-excel-cells.html'),
         new menuItem('Rich text format in cell',folderParentFile + '../../basics/excel-rich-text-format.html'),
         new menuItem('Merge cells',folderParentFile + '../../basics/excel-merge-cells.html'),
         new menuItem('Column width and row height',folderParentFile + '../../basics/excel-column-width-row-height.html'),
         mnuAutoformatting,
         new menuItem('Conditional formatting',folderParentFile + '../../basics/excel-conditional-formatting.html'),
         new menuItem('Themes',folderParentFile + '../../basics/excel-theme.html'));
      
       var mnuPageSetup = new menu('mnuPageSetup', 
         new menuTrigger('trgPageSetup','Page setup'),
         new menuItem('Page setup',folderParentFile + '../../basics/excel-page-setup.html'),       
         new menuItem('Header and footer',folderParentFile + '../../basics/excel-header-footer.html'),
         new menuItem('Page breaks',folderParentFile + '../../basics/excel-sheet-page-breaks.html')); 
         
       var mnuCharts = new menu('mnuCharts', 
         new menuTrigger('trgCharts','Charts'),
         new menuItem('Chart inside a worksheet',folderParentFile + '../../basics/excel-chart-inside-sheet.html'), 
         new menuItem('Chart sheet',folderParentFile + '../../basics/excel-chart-sheet.html'),
         new menuItem('Chart types',folderParentFile + '../../basics/excel-chart-types.html'),    
         new menuItem('Series',folderParentFile + '../../basics/excel-chart-series.html'),
         new menuItem('Category X axis',folderParentFile + '../../basics/excel-chart-category-x-axis.html'),
         new menuItem('Value Y axis',folderParentFile + '../../basics/excel-chart-value-y-axis.html'),
         new menuItem('Chart area',folderParentFile + '../../basics/excel-chart-area.html'),
         new menuItem('Plot area',folderParentFile + '../../basics/excel-chart-plot-area.html'),
         new menuItem('Legend',folderParentFile + '../../basics/excel-chart-legend.html'),
         new menuItem('Gridlines',folderParentFile + '../../basics/excel-chart-gridlines.html'),
         new menuItem('Data table',folderParentFile + '../../basics/excel-chart-data-table.html'),
         new menuItem('Chart title and axis titles',folderParentFile + '../../basics/excel-chart-axis-title.html'), 
         new menuItem('3D rotation', folderParentFile + '../../basics/excel-chart-3d-rotation.html')); 
         
       var mnuPivotTables = new menu('mnuPivotTables', 
         new menuTrigger('trgPivotTables','Pivot tables and pivot charts'),
         new menuItem('Pivot tables',folderParentFile + '../../basics/excel-pivot-table.html'), 
         new menuItem('Pivot charts',folderParentFile + '../../basics/excel-pivot-chart.html')); 
         
       var mnuSecurityAndProtection = new menu('mnuSecurityAndProtection', 
         new menuTrigger('trgmnuSecurityAndProtection','Security and protection'),      
         new menuItem('Protect sheet elements',folderParentFile + '../../basics/excel-protect-sheet.html'),
         new menuItem('Password protected Excel file and encryption', folderParentFile + '../../basics/password-protected-excel-file.html'));

       var mnuConvertFiles = new menu('mnuConvertFiles',
         new menuTrigger('trgmnuConvertFiles', 'Convert files'),
         new menuItem('Convert HTML file to Excel', folderParentFile + '../../basics/convert-html-to-excel.html'),
         new menuItem('Convert CSV file to Excel', folderParentFile + '../../basics/convert-csv-to-excel.html'),
         new menuItem('Convert XML file to Excel', folderParentFile + '../../basics/convert-xml-to-excel.html'),
         new menuItem('Convert Excel file to HTML', folderParentFile + '../../basics/convert-excel-to-html.html'),
         new menuItem('Convert Excel file to CSV', folderParentFile + '../../basics/convert-excel-to-csv.html'),
         new menuItem('Convert Excel file to XML', folderParentFile + '../../basics/convert-excel-to-xml.html')); 
      
      var mnuEasyXLSBasics = new menu('mnuEasyXLSBasics',
         new menuTrigger('trgEasyXLSBasics','EasyXLS<sup>TM</sup> Basics'),
         //new menuItem('Easy XLS Basics','../html/programmerGuide/basics/easyXLSbasics.htm'), 
         new menuItem('Create an Excel document',folderParentFile + '../../basics/create-excel-file.html'),
         mnuImportFiles,
	     mnuExportFiles,
	     mnuConvertFiles,
         mnuSheets,
         mnuAddressingCells,
         mnuCellFormatting,
         new menuItem('Cell comments',folderParentFile + '../../basics/excel-cell-comment.html'),
         new menuItem('Hyperlinks',folderParentFile + '../../basics/excel-hyperlink.html'),
         new menuItem('Images',folderParentFile + '../../basics/excel-image-import-export.html'),
         new menuItem('Named ranges and formulas',folderParentFile + '../../basics/excel-named-range-and-formula.html'),
         new menuItem('Data validation',folderParentFile + '../../basics/excel-data-validation.html'),
         mnuPageSetup,
         new menuItem('Macros and VBA project', folderParentFile + '../../basics/excel-macros-vba-project.html'),
		 new menuItem('Groups and outline levels', folderParentFile + '../../basics/excel-group-outline-levels.html'),
         new menuItem('Freeze and split panes', folderParentFile + '../../basics/excel-split-freeze-pane.html'),
		 new menuItem('Filter and autofilter',folderParentFile + '../../basics/excel-filter-and-autofilter.html'),
         mnuCharts,
         mnuPivotTables,
         mnuSecurityAndProtection,
         new menuItem('Document properties',folderParentFile + '../../basics/excel-document-properties.html')
         //new menuItem('Load CSV File','../html/programmerGuide/basics/loadCSVfile.html') ,
    	 //new menuItem('Export CSV File','../html/programmerGuide/basics/exportCSVFile.html')	
         );
      mnuEasyXLSBasics.position(-1,10,230);
      mnuEasyXLSBasics.write(); 

       var mnuFAQ = new menu('mnuFAQ',
         new menuTrigger('trgFAQ','FAQ'),
         new menuItem('How to export to Excel file in C# and VB.NET', folderParentFile + '../../FAQ/export-to-excel-in-dot-net.html'),
         new menuItem('How to export DataTable to Excel file in .NET', folderParentFile + '../../FAQ/export-datatable-to-excel.html'),
         new menuItem('How to export GridView to Excel file in ASP.NET',folderParentFile + '../../FAQ/export-gridview-to-excel.html'),
         new menuItem('How to export DataGridView to Excel file in .NET',folderParentFile + '../../FAQ/export-datagridview-to-excel.html'),
         new menuItem('How to export DataGrid to Excel file in .NET',folderParentFile + '../../FAQ/export-datagrid-to-excel.html'),
         new menuItem('How to import Excel file in C# and VB.NET', folderParentFile + '../../FAQ/import-excel-in-dot-net.html'),
         new menuItem('How to import Excel file to SQL table in .NET',folderParentFile + '../../FAQ/import-excel-to-sql.html'),
         new menuItem('How to import Excel file to DataTable in .NET',folderParentFile + '../../FAQ/import-excel-to-datatable.html'),
         new menuItem('How to import Excel file to GridView in .NET',folderParentFile + '../../FAQ/import-excel-to-gridview.html'),
         new menuItem('How to import Excel file to DataGridView in .NET',folderParentFile + '../../FAQ/import-excel-to-datagridview.html'),
         new menuItem('How to import Excel file to DataGrid in .NET',folderParentFile + '../../FAQ/import-excel-to-datagrid.html'),
         new menuItem('How to read Excel file in C# and VB.NET',folderParentFile + '../../FAQ/read-excel-file-in-dot-net.html'),
         new menuItem('How to export to Excel file in PHP and ASP classic', folderParentFile + '../../FAQ/export-to-excel-in-php-asp.html'),
         new menuItem('How to import Excel file in PHP   and ASP classic', folderParentFile + '../../FAQ/import-excel-in-php-asp.html'),
         new menuItem('How to import Excel data to MySQL, SQL Server in PHP or ASP classic', folderParentFile + '../../FAQ/import-excel-to-mysql.html')
         );

      mnuFAQ.position(-1,10,230);
      mnuFAQ.write();
      
      var mnuTipsAndTricks = new menu('mnuTipsAndTricks',
         new menuTrigger('trgTipsAndTricks','Tips and Tricks'),
		 new menuItem('Export large Excel files', folderParentFile + '../../tips-and-tricks/export-large-excel-file.html'),
		 new menuItem('Read large Excel files', folderParentFile + '../../tips-and-tricks/read-large-excel-file.html'),
		 new menuItem('How to install EasyXLS for Java on any operating system', folderParentFile + '../../tips-and-tricks/install-easyxls-for-java.html'),
		 new menuItem('How to write and open an Excel file in browser using Response stream', folderParentFile + '../../tips-and-tricks/open-excel-in-browser-response-stream.html'),
		 new menuItem('How to create drop down list in Excel',folderParentFile + '../../tips-and-tricks/excel-drop-down-list.html'),
		 new menuItem('How to retrieve, display or hide error messages', folderParentFile + '../../tips-and-tricks/error-messages.html')
		 );
		 
	  mnuTipsAndTricks.position(-1,10,230);
      mnuTipsAndTricks.write(); 
     
	   var mnuTroubleshooting = new menu('mnuTroubleshooting',
         new menuTrigger('trgTroubleshooting', 'Troubleshooting'),
         new menuItem('Installation', folderParentFile + '../../troubleshooting/easyxls-installation-errors.html'),
         new menuItem('Invalid license', folderParentFile + '../../troubleshooting/easyxls-licensing-errors.html'),
         new menuItem('.NET runtime errors', folderParentFile + '../../troubleshooting/easyxls-dotnet-runtime-errors.html'),
         new menuItem('PHP runtime errors', folderParentFile + '../../troubleshooting/easyxls-php-runtime-errors.html'),
         new menuItem('VBScript runtime errors', folderParentFile + '../../troubleshooting/easyxls-vbscript-runtime-errors.html'),
         new menuItem('C++ runtime errors', folderParentFile + '../../troubleshooting/easyxls-cpp-runtime-errors.html'),
         new menuItem('Java runtime errors', folderParentFile + '../../troubleshooting/easyxls-java-runtime-errors.html')
		 );
		 
      mnuTroubleshooting.position(-1,10,230);
      mnuTroubleshooting.write(); 	

	   
		 
		 
	  
	  fin();  
	  
	  }