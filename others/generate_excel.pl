use Win32::OLE;
# Start Excel and make it visible
$xlApp = Win32::OLE->new('Excel.Application');
$xlApp->{Visible} = 0;

# Create a new workbook
$xlBook = $xlApp->Workbooks->Add;

# Our data that we will add to the workbook...
$mydata = [["Item",     "Category", "Price"], 
           ["Nails",    "Hardware",  "5.25"],
           ["Shirt",    "Clothing", "23.00"],
           ["Hammer",   "Hardware", "16.25"],
           ["Sandwich", "Food",      "5.00"],
           ["Pants",    "Clothing", "31.00"],
           ["Drinks",   "Food",      "2.25"]];

# Write all the data at once...
$rng = $xlBook->workSheets(1)->Range("A1:C7");
$rng->{Value} = $mydata;

# Wait for user input...
print "Press <return> to continue...";
$x = <STDIN>;

# Clean up

$xlBook->saveas('C:\Users\jayhsieh\Desktop\abc.xlsx');
$xlApp->Quit;
$xlBook = 0;
$xlApp = 0;

print "All done.";