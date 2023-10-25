#include <iostream>
#include <xlnt/xlnt.hpp>

using namespace std;

int main()
{
    // Load an existing Excel file
    xlnt::workbook wb;
    wb.load("excel_file3.xlsx");

    // Select the active worksheet
    xlnt::worksheet ws = wb.active_sheet();

    //going through cells
    for (auto row : ws.rows(false))
    {
        for (auto cell : row)
        {
            //if value of cell is <=2 then fill it with darkgrenn
            //note this commend fill also empty cells bc they are treated as 0  
            if (cell.value<int>() <= 2)
            {
                cell.fill(xlnt::fill::solid(xlnt::color::darkgreen()));
            }
        }
    }

    wb.save("excel_file3.xlsx");

    clog << "Processing complete" << endl;
   

    return 0;
}
