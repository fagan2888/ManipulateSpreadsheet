#include <stdio.h>
#include <sal/main.h>
#include <cppuhelper/bootstrap.hxx>
#include <com/sun/star/lang/XMultiComponentFactory.hpp>
#include <com/sun/star/frame/XDesktop2.hpp>
#include <com/sun/star/sheet/XSpreadsheetDocument.hpp>
#include <com/sun/star/sheet/XSpreadsheet.hpp>
#include <com/sun/star/sheet/XSpreadsheetView.hpp>
#include <com/sun/star/sheet/XCellRangesQuery.hpp>
#include <com/sun/star/sheet/XSheetCellRanges.hpp>
#include <com/sun/star/sheet/XCellAddressable.hpp>
#include <com/sun/star/sheet/CellFlags.hpp>
#include <com/sun/star/table/XCell.hpp>
#include <com/sun/star/table/CellVertJustify2.hpp>
#include <com/sun/star/beans/XPropertySet.hpp>
#include <com/sun/star/frame/XModel.hpp>
#include <com/sun/star/frame/XController.hpp>
#include <com/sun/star/container/XEnumerationAccess.hpp>
#include <com/sun/star/container/XEnumeration.hpp>


using namespace com::sun::star::uno;
using namespace com::sun::star::lang;
using namespace com::sun::star::frame;
using namespace com::sun::star::sheet;
using namespace com::sun::star::beans;
using namespace com::sun::star::table;
using namespace com::sun::star::container;


using ::rtl::OString;
using ::rtl::OUString;
using ::rtl::OUStringToOString;

SAL_IMPLEMENT_MAIN()
{
    try
    {
        // get the remote office component context
        Reference< XComponentContext > xContext( ::cppu::bootstrap() );
	if ( !xContext.is() )
	{
	    fprintf(stdout, "\nError getting context from running LO instance...\n");
	    return -1;
	}
	
	// retrieve the service-manager from the context
        Reference< XMultiComponentFactory > xServiceManager = xContext->getServiceManager();
	if ( xServiceManager.is() )
	    fprintf(stdout, "\nremote ServiceManager is available\n");
	else
	    fprintf(stdout, "\nremote ServiceManager is not available\n");
	fflush(stdout);

	// Create Desktop object, xDesktop is still a XInterface, need to downcast it.
	Reference< XInterface > xDesktop = xServiceManager->createInstanceWithContext( OUString("com.sun.star.frame.Desktop"), xContext );

	if ( !xDesktop.is() )
	{
	    fprintf(stdout, "\nError creating com.sun.star.frame.Desktop object");
	    fflush(stdout);
	    return -1;
	}

	// Get XDesktop2 from XInterface type indirectly using XInterface's queryInterface() method.
	Reference< XDesktop2 > xDesktop2( xDesktop, UNO_QUERY );

	if ( !xDesktop2.is() )
	{
	    fprintf(stdout, "\nError upcasting XInterface to XDesktop2");
	    fflush(stdout);
	    return -1;
	}

	// open a spreadsheet document                                                                                                                                                                 
        Reference< XComponent > xComponent = xDesktop2->loadComponentFromURL( OUString( "private:factory/scalc" ), // URL to the ods file
									      OUString( "_blank" ), 0,
									      Sequence < ::com::sun::star::beans::PropertyValue >() );
	if ( !xComponent.is() )
	{
	    fprintf(stdout, "\nopening spreadsheet document failed!\n");
	    fflush(stdout);
	    return -1;
	}
	
	fprintf(stdout, "\nopened spreadsheet document...\n");
	fflush(stdout);

	// Get XSpreadsheetDocument interface
	Reference< XSpreadsheetDocument > xSpreadsheetDocument(xComponent, UNO_QUERY);

	// Get object that supports XSpreadsheets which represents a collection of XSpreadsheet
	Reference< XSpreadsheets > xSpreadsheets = xSpreadsheetDocument->getSheets();
	xSpreadsheets->insertNewByName( OUString("MySheet"), (short)0 );

	// Get type of xSpreadsheets's element type using the interface XElementAccess which XSpreadsheets inherits.
	// Returns void type if xSpreadsheets is a multi-type container.
	Type aElemType = xSpreadsheets->getElementType();

	fprintf( stdout, "\n>>> xSpreadsheets container has elements each of type = %s\n",
		 OUStringToOString( aElemType.getTypeName(), RTL_TEXTENCODING_ASCII_US ).getStr() );
	fflush(stdout);

	// Get MySheet sheet object from the collection xSpreadsheets (via interface XNameAccess)
	Any aSheet = xSpreadsheets->getByName( OUString("MySheet") );
	// Convert Any object to its real type - XSpreadsheet
	Reference< XSpreadsheet > xSpreadsheet( aSheet, UNO_QUERY );

	// Get A1 cell object (Cell service) defined by the interface XCell using
	// getCellByPosition method of XCellRange interface inherited XSpreadsheet
	// getCellByPosition takes column idx and row idx in order.
	Reference< XCell > xCell = xSpreadsheet->getCellByPosition(0, 0);
	xCell->setValue(105);

	xCell = xSpreadsheet->getCellByPosition(0, 1);
	xCell->setValue(501);
	
	xCell = xSpreadsheet->getCellByPosition(0, 2);
	xCell->setFormula("=sum(A1:A2)");

	// Get XPropertySet interface of xCell object which is really a Cell service object
	Reference< XPropertySet > xCellProps( xCell, UNO_QUERY );
	// Set property "CellStyle" to the value "Result"
	xCellProps->setPropertyValue( OUString("CellStyle"), makeAny( OUString("Result") ) );

	// Set property using an enum
	xCellProps->setPropertyValue( OUString("VertJustify"), makeAny( CellVertJustify2::TOP ) );

	Reference< XModel > xSpreadsheetModel( xComponent, UNO_QUERY );
	Reference< XController > xSpreadsheetController = xSpreadsheetModel->getCurrentController();
	Reference< XSpreadsheetView > xSpreadsheetView( xSpreadsheetController, UNO_QUERY );

	// Make the new sheet the active sheet
	xSpreadsheetView->setActiveSheet( xSpreadsheet );

	// Get XCellRangesQuery supporting object
	Reference< XCellRangesQuery > xCellQuery( aSheet, UNO_QUERY );

	// Get the cell ranges where there are formulas in aSheet
	Reference< XSheetCellRanges > xFormulaCells = xCellQuery->queryContentCells( (short)CellFlags::FORMULA );

	// Get collection of all formula cells
	Reference< XEnumerationAccess > xFormulasEnumAccess = xFormulaCells->getCells();
	Reference< XEnumeration > xFormulaEnum = xFormulasEnumAccess->createEnumeration();

	while ( xFormulaEnum->hasMoreElements() )
	{
	    Any aNextElement = xFormulaEnum->nextElement();
	    xCell = Reference< XCell >( aNextElement, UNO_QUERY );
	    Reference< XCellAddressable > xCellAddress( xCell, UNO_QUERY );
	    fprintf( stdout, ">>> Formula cell in column = %d, row = %d, with formula = %s\n",
		     xCellAddress->getCellAddress().Column,
		     xCellAddress->getCellAddress().Row,
		     OUStringToOString( xCell->getFormula(), RTL_TEXTENCODING_ASCII_US ).getStr() );
	    fflush( stdout );
	}
	
    }
    catch ( ::cppu::BootstrapException& e )
    {
        fprintf(stderr, "caught BootstrapException: %s\n",
                OUStringToOString( e.getMessage(), RTL_TEXTENCODING_ASCII_US ).getStr());
	fflush(stderr);
        return 1;
    }
    catch ( Exception& e )
    {
        fprintf(stderr, "caught UNO exception: %s\n",
                OUStringToOString( e.Message, RTL_TEXTENCODING_ASCII_US ).getStr());
	fflush(stderr);
        return 1;
    }

    return 0;

}
