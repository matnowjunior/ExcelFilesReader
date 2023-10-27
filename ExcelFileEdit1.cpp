#include <string>
#include <iostream>
#include <xlnt/xlnt.hpp>
#include <fstream>

#define RED "\033[31m"    
#define GREEN "\033[32m" 
#define RESET "\033[0m"

using namespace std;

int row_number1, col_number1, row_number2, col_number2;//global variables

pair <string, string> CopyFile(string originalFileName)
{
    string newFileName = "excel_new.xlsx";  //new file name
    string announcement;

    ifstream source(originalFileName, ios::binary); //opening file in binary mode
    ofstream dest(newFileName, ios::binary); //saving file to new one   

    if (source && dest) { //checking if both streams were opened correctly
        dest << source.rdbuf(); //coping the content of the source file to the target file
        source.close();//closing soource file
        dest.close();//closing target file
        announcement ="Excel file has been sucessfully copied";
    }
    else {
        announcement = "An error occurred while copying and renaming the file.";
    }
    return make_pair(newFileName, announcement);
}

int titleToNumber(string s)
{
    int r = 0;
    for (int i = 0; i < s.length(); i++)
    {
        r = r * 26 + s[i] - 64;
    }
    return r;
}

pair <int, int> signs_numbers_separately(string cell_input)
{
    string letters, digits;
    for (char c : cell_input) {
        if (isalpha(c)) {
            letters += c;  //adding letter to letter variable
        }
        else if (isdigit(c)) {
            digits += c;  //adding number to digits variable
        }
    }

    return make_pair(titleToNumber(letters), stoi(digits));

}

int main()
{
    string originalFileName;
    cout << "Provide the excel file path: " << endl;
    cin >> originalFileName;

    //Opening Excel workbook and worksheet
    xlnt::workbook wb;
    xlnt::worksheet ws;

    //Loading excel file
    try {
        wb.load(CopyFile(originalFileName).first);  // excel file path
        ws = wb.active_sheet();//getting active worksheet from workbook
    }
    catch (const xlnt::exception& e) {
        cout << RED << "Processing failed " << RESET << e.what() << std::endl;//displaying an error message when trying to read a file
        return 1;
    }

    string cell_input1, cell_input2;

    cout << "Podaj pierwsza komorke:";
    cin >> cell_input1;


    col_number1 = signs_numbers_separately(cell_input1).first;
    row_number1 = signs_numbers_separately(cell_input1).second;

    cout << "Podaj druga komorke:";
    cin >> cell_input2;

    col_number2 = signs_numbers_separately(cell_input2).first;
    row_number2 = signs_numbers_separately(cell_input2).second;


    //looping through (before) specified range
    for (int i = row_number1; i < row_number2 + 1; i++)
    {

        for (int j = col_number1; j < col_number2 + 1; j++)
        {

            //setting a new cell to the entered values(col and row)
            xlnt::cell cell = ws.cell(i, j);
            cout << cell.value<float>();
            
            //cell.clear_style();
           
            if(cell.value<float>() < 30)
                cell.fill(xlnt::fill::solid(xlnt::rgb_color(242, 34, 15)));
            else if(cell.value<float>() >=30 && cell.value<float>() < 50)
                cell.fill(xlnt::fill::solid(xlnt::rgb_color(242, 160, 7)));
            else if(cell.value<float>() >=50 && cell.value<float>() < 75)
                cell.fill(xlnt::fill::solid(xlnt::rgb_color(5, 242, 108))); 
            else if(cell.value<float>() >= 75 && cell.value<float>() < 90)
                cell.fill(xlnt::fill::solid(xlnt::rgb_color(5, 151, 242)));
            else if(cell.value<float>() >= 90 && cell.value<float>() <100 )
                cell.fill(xlnt::fill::solid(xlnt::rgb_color(178, 37, 217)));
            cout << " ";
        }
        cout << endl;
    }
        
    

    wb.save(CopyFile(originalFileName).first);
    CopyFile(originalFileName).second;

    cout << GREEN << "Processing succeed :)" << RESET;
}
