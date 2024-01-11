#include <iostream>
#include <vector>
#include <iomanip>
#include <string>
#include <conio.h>
#include <fstream>
#include <sstream>
#include <thread>
#include <chrono>



#ifdef _WIN32
#include <windows.h>
#else
#include <unistd.h>
#endif

using namespace std;
HANDLE acolor = GetStdHandle(STD_OUTPUT_HANDLE);

int MAX_ROWS = 1;
int MAX_COLS = 1;



// Function to print a loading bar
void printLoadingBar(int progress, int total, int barWidth)
{
    float percentage = static_cast<float>(progress) / total;
    int numChars = static_cast<int>(percentage * barWidth);

    // Print the loading bar
    std::cout << "[";
    for (int i = 0; i < barWidth; ++i) {
        if (i < numChars) {
            // Print colored character for the progress
#ifdef _WIN32
            SetConsoleTextAttribute(GetStdHandle(STD_OUTPUT_HANDLE), 2);  // 2 is the ANSI color code for green
#else
            std::cout << "\033[32m";
#endif
            std::cout << "=";
        }
        else {
            std::cout << " ";
        }
    }
#ifndef _WIN32
    std::cout << "\033[0m";  // Reset color on Unix-like systems
#endif
    std::cout << "] " << std::setprecision(3) << percentage * 100.0 << "%\r";
    std::cout.flush();
}



template<typename T>
class MiniExcelClass
{
private:
    vector<vector<string>> grid;
    vector<T>clipboardVector;
    int currentRow = 1;
    int currentCol = 1;
public:
    template<typename T>
    class NodeCell {
    public:

        T value;
        NodeCell* up;
        NodeCell* down;
        NodeCell* left;
        NodeCell* right;

        NodeCell()
        {

        }
        NodeCell(T val)
        {
            value = val;
            up = nullptr;
            down = nullptr;
            left = nullptr;
            right = nullptr;
        }

    };

    class Iterator {
    private:
        NodeCell<T>* current;
    public:
        Iterator(NodeCell<T>* iter)
        {
            current = iter;
        }

        Iterator& operator++()
        {
            if (current->down != nullptr)
                current = current->down;
            return *this;
        }


        Iterator& operator--()
        {
            if (current->up != nullptr)
                current = current->up;
            return *this;
        }

        Iterator operator++(int)
        {
            Iterator temp = *this;
            if (current->right != nullptr)
                current = current->right;
            return temp;
        }

        Iterator operator--(int)
        {
            Iterator temp = *this;
            if (current->left != nullptr)
                current = current->left;
            return temp;
        }
        bool operator==(Iterator newIter)
        {
            return current == newIter.current;
        }

        bool operator!=(Iterator newIter)
        {
            return current != newIter.current;
        }

        T& operator * ()
        {
            return current->value;
        }
        NodeCell<T>* get()
        {
            return current;
        }

    };

    Iterator begin()
    {
        return Iterator(current);
    }

    Iterator end()
    {
        return Iterator(nullptr);
    }

    NodeCell<T>* current;

    MiniExcelClass() : grid(MAX_ROWS, std::vector<std::string>(MAX_COLS, "")), currentRow(0), currentCol(0)
    {
        current = new NodeCell<T>();
    }

    ~MiniExcelClass() {

    }


    void moveUp() {
        if (currentRow > 0) {
            currentRow--;
            current = current->up;
        }
    }

    void moveDown() {
        if (currentRow < MAX_ROWS - 1) {
            currentRow++;
            current = current->down;

        }
    }

    void moveLeft() {
        if (currentCol > 0) {
            currentCol--;
            current = current->left;
        }
    }

    void moveRight() {
        if (currentCol < MAX_COLS - 1) {
            currentCol++;
            current = current->right;

        }
    }


   /* void displaySheet() {
    
        system("cls");

        std::cout << "Excel Sheet:" << std::endl;
        for (int i = 0; i < MAX_ROWS; ++i) {
            for (int j = 0; j < MAX_COLS; ++j) {
                std::cout << "+-------------------";
            }
            std::cout << "+" << std::endl;

            for (int j = 0; j < MAX_COLS; ++j) {
                std::cout << "| " << std::setw(17) << std::left << grid[i][j] << " ";
            }
            std::cout << "|" << std::endl;
        }

        for (int j = 0; j < MAX_COLS; ++j) {
            std::cout << "+-------------------";
        }
        std::cout << "+" << std::endl;

        std::cout << "Current Cell: Row " << currentRow + 1 << ", Column " << currentCol + 1 << std::endl;
    }*/

    void displaySheet() {
        system("cls");

        HANDLE hConsole = GetStdHandle(STD_OUTPUT_HANDLE);
        std::cout << "EXCEL SHEET:" << std::endl << endl;
        for (int i = 0; i < MAX_ROWS; ++i) {
            for (int j = 0; j < MAX_COLS; ++j) {
                if (i == currentRow && j == currentCol) {
                    // Set text color to blue for the current cell
                    SetConsoleTextAttribute(hConsole, FOREGROUND_RED);
                }
                else {
                    // Set text color to white for other cells
                    SetConsoleTextAttribute(hConsole, FOREGROUND_RED | FOREGROUND_GREEN | FOREGROUND_BLUE);
                }

                std::cout << "+-------------------";
            }
            std::cout << "+" << std::endl;

            for (int j = 0; j < MAX_COLS; ++j) {
                if (i == currentRow && j == currentCol) {
                    // Set text color to blue for the current cell
                    SetConsoleTextAttribute(hConsole, FOREGROUND_RED);
                }
                else {
                    // Set text color to white for other cells
                    SetConsoleTextAttribute(hConsole, FOREGROUND_RED | FOREGROUND_GREEN | FOREGROUND_BLUE);
                }

                std::cout << "| " << std::setw(17) << std::left << grid[i][j] << " ";
            }
            std::cout << "|" << std::endl;
        }

        for (int j = 0; j < MAX_COLS; ++j) {
            // Set text color to white for the last row
            SetConsoleTextAttribute(hConsole, FOREGROUND_RED | FOREGROUND_GREEN | FOREGROUND_BLUE);
            std::cout << "+-------------------";
        }
        std::cout << "+" << std::endl;

        // Reset text color to default
        SetConsoleTextAttribute(hConsole, FOREGROUND_RED | FOREGROUND_GREEN | FOREGROUND_BLUE);

        std::cout << "Current Cell: Row " << currentRow + 1 << ", Column " << currentCol + 1 << std::endl;
    }
   


    void writeData(const std::string& data) {
        if (currentRow >= 0 && currentRow < MAX_ROWS && currentCol >= 0 && currentCol < MAX_COLS) {
            grid[currentRow][currentCol] = data;
        }
        else {
            std::cout << "No current cell selected." << std::endl;
        }
    }


    NodeCell<T>* getCurrentNode() const
    {
        return current;
    }



    void InsertRowBelow()
    {
        NodeCell<T>* temp = current;
        if (temp == nullptr)
        {
            current = new NodeCell<T>();
        }
        else if (temp->down == nullptr)
        {
            while (temp->down != nullptr)
            {
                temp = temp->down;
            }
            while (temp->left != nullptr)
            {
                temp = temp->left;
            }

            temp->down = new NodeCell<T>();
            temp->down->up = temp;
            temp->down->right = nullptr;
            temp->down->down = nullptr;


            while (temp->right != nullptr)
            {
                NodeCell<T>* newCell = new NodeCell<T>();
                temp->down->right = newCell;
                newCell->left = temp->down;

                temp = temp->right;
                newCell->up = temp;
                temp->down = newCell;
                newCell->down = nullptr;

            }

            grid.emplace_back(MAX_COLS, "");
            MAX_ROWS++;

            //current = temp->down;
        }
        else if (temp->down != nullptr)
        {
            while (temp->left != nullptr)
            {
                temp = temp->left;
            }
            
            NodeCell<T>* newNode = new NodeCell<T>();
            newNode->up = temp;
            newNode->down = temp->down;
            temp->down->up = newNode;
            temp->down = newNode;
            
            if (temp->right != nullptr)
            {
                temp = temp->right;
                while (temp != nullptr)
                {
                    NodeCell<T>* newNode1 = new NodeCell<T>();
                    newNode1->up = temp;
                    newNode1->down = temp->down;
                    temp->down->up = newNode1;
                    temp->down = newNode1;
                    newNode1->left = temp->left->down;
                    temp->left->down->right = newNode1;
                    temp = temp->right;
                }
            }
            grid.insert(grid.begin() + currentRow + 1, std::vector<std::string>(MAX_COLS, ""));
            MAX_ROWS++;
            
        }
    }

    void insertCoulmnToRight()
    {
        NodeCell<T>* temp = current;
        if (temp == nullptr)
        {
            current = new NodeCell<T>();
        }
        else if (temp->right == nullptr)
        {

            while (temp->up != nullptr)
            {
                temp = temp->up;
            }

            temp->right = new NodeCell<T>();
            temp->right->left = temp;
            temp->right->down = nullptr;

            while (temp->down != nullptr)
            {
                NodeCell<T>* newNode = new NodeCell<T>();
                temp = temp->down;
                temp->right = newNode;
                newNode->left = temp;
                newNode->up = temp->up->right;
                temp->up->right->down = newNode;
                newNode->right = nullptr;

            }
           

            for (int i = 0; i < MAX_ROWS; ++i)
            {
                grid[i].emplace_back("");
            }

            MAX_COLS++;
            //current = temp->left;

        }
        else
        {
            while (temp->up != nullptr)
            {
                temp = temp->up;
            }

            NodeCell<T>* newNode1 = new NodeCell<T>();
            newNode1->left = temp;
            newNode1->right = temp->right;
            temp->right->left = newNode1;
            temp->right = newNode1;
            temp = temp->down;

            while (temp != nullptr)
            {
                NodeCell<T>* newNode = new NodeCell<T>();
                newNode->left = temp;
                newNode->right = temp->right;
                temp->right->left = newNode;
                temp->right = newNode;
                newNode->up = temp->up->right;
                temp->up->right->down = newNode;
                temp = temp->down;
            }
            for (int i = 0; i < MAX_ROWS; ++i)
            {
                grid[i].insert(grid[i].begin() + currentCol + 1, ""); // Insert an empty string at the correct position
            }

            MAX_COLS++;
            
        }
    }


    void insertRowAbove()
    {
        NodeCell<T>* temp = current;
        if (temp == nullptr)
        {
            current = new NodeCell<T>();
        }
        else if(temp->up == nullptr)
        {
            while (temp->left != nullptr)
            {
                temp = temp->left;
            }
            temp->up = new NodeCell<T>();
            temp->up->down = temp;

            // Update grid
            grid.insert(grid.begin(), std::vector<std::string>(MAX_COLS, ""));

            while (temp->right != nullptr)
            {
                NodeCell<T>* newNode = new NodeCell<T>();
                newNode->left = temp->up;
                temp->up->right = newNode;
                temp = temp->right;
                temp->up = newNode;
                newNode->down = temp;

            }
            grid.emplace_back(MAX_COLS, "");
            MAX_ROWS++;
           current = current->up;
        }
        else if (temp->up != nullptr)
        {
            grid.insert(grid.begin() + currentRow, std::vector<std::string>(MAX_COLS, ""));
            while (temp->left != nullptr)
            {
                temp = temp->left;
            }
            NodeCell<T>* newNode = new NodeCell<T>();
            newNode->down = temp;
            newNode->up = temp->up;
            temp->up->down = newNode;
            temp->up = newNode;
           
            if (temp->right != nullptr)
            {
                temp = temp->right;
                while (temp->right != nullptr)
                {
                    NodeCell<T>* node1 = new NodeCell<T>();
                    node1->down = temp;
                    node1->up = temp->up;
                    temp->up->down = node1;
                    temp->up = node1;
                    node1->left = temp->left->up;
                    temp->left->up->right = node1;
                    temp = temp->right;
                }
            }
            grid.emplace_back(MAX_COLS, "");
            MAX_ROWS++;
            current = current->up;
        }
        
    }

    void insertColumntoLeft()
    {
        NodeCell<T>* temp = current;
        if (temp == nullptr)
        {
            current = new NodeCell<T>();
        }
        else if (temp->left == nullptr)
        {
            while (temp->up != nullptr)
            {
                temp = temp->up;
            }
            temp->left = new NodeCell<T>();
            temp->left->right = temp;

            // Update grid
            for (int i = 0; i < MAX_ROWS; ++i)
            {
                grid[i].insert(grid[i].begin(), "");
            }


            while (temp->down != nullptr)
            {
                NodeCell<T>* newNode = new NodeCell<T>();
                newNode->up = temp->left;
                temp->left->down = newNode;
                temp = temp->down;
                newNode->right = temp;
                temp->left = newNode;
            }
            for (int i = 0; i < MAX_ROWS; ++i)
            {
                grid[i].emplace_back("");
            }

            MAX_COLS++;
            current = current->left;
        }
        else if (temp->left != nullptr)
        {
            while (temp->up != nullptr)
            {
                temp = temp->up;
            }
            NodeCell<T>* col = new NodeCell<T>();

            col->left = temp->left;
            col->right = temp;
            temp->left->right = col;
            temp->left = col;
            
            if (temp->down != nullptr)
            {
                temp = temp->down;
                while (temp->down != nullptr)
                {
                    NodeCell<T>* newNode = new NodeCell<T>();
                    newNode->left = temp->left;
                    newNode->right = temp;
                    temp->left->right = newNode;
                    temp->left = newNode;
                    newNode->up = temp->up->left;
                    temp->up->left->down = newNode;
                    temp = temp->down;
                }
            }
            
            for (int i = 0; i < MAX_ROWS; ++i)
            {
                grid[i].insert(grid[i].begin() + currentCol, ""); // Insert an empty string at the correct position
            }

            MAX_COLS++;
            current = current->left;

        }
    }
            
   
    void insertCellByRightShift()
    {

        NodeCell<T>* temp = current;
        if (current == nullptr) 
        {
            return;
        }
        while (current->right != nullptr)
        {
            current = current->right;
        }
        insertCoulmnToRight();
        current = current->right;

        while (temp != current)
        {
            current->value = current->left->value;
            current = current->left;
        }
        temp->value = "";

        //FOR UPDATING THE GRID

        for (int i = MAX_COLS - 2; i >= currentCol; --i)
        {
            grid[currentRow][i + 1] = grid[currentRow][i];
        }

        grid[currentRow][currentCol] = "";
      
    }

    void insertCellByDownShift()
    {
        NodeCell<T>* temp = current;
        if (current == nullptr)
        {
            return;
        }
        while (current->down != nullptr)
        {
            current = current->down;
        }
        InsertRowBelow();
        current = current->down;
        while (temp != current)
        {
            current->value = current->up->value;
            current = current->up;
        }
        temp->value = "";

        //FOR UPDATING THE GRID

        for (int i = MAX_ROWS - 2; i >= currentRow; --i)
        {
            grid[i + 1][currentCol] = grid[i][currentCol];
        }

        grid[currentRow][currentCol] = "";
    }

    void deleteCellbyLeftShift()
    {
        NodeCell<T>* temp = current;
        temp->value = "";
        while (temp->right != nullptr)
        {
            temp->value = temp->right->value;
            temp = temp->right;
        }
        temp->value = "";
        //FOR UPDATING THE GRID
        for (int i = currentCol; i < MAX_COLS - 1; ++i)
        {
            grid[currentRow][i] = grid[currentRow][i + 1];
        }

        grid[currentRow][MAX_COLS - 1] = "";
       
    }

    void deleteCellbyUpShift()
    {
        NodeCell<T>* temp = current;
        temp->value = "";
        while (temp->down != nullptr)
        {
            temp->value = temp->down->value;
            temp = temp->down;
        }
        temp->value = "";
        //FOR UPDATING THE GRID
        for (int i = currentRow; i < MAX_ROWS - 1; ++i)
        {
            grid[i][currentCol] = grid[i + 1][currentCol];
        }

        grid[MAX_ROWS - 1][currentCol] = "";
    }

    void clearRow()
    {
        NodeCell<T>* temp = current;
        while (temp->left != nullptr)
        {
            temp = temp->left;
        }
        while (temp != nullptr)
        {
            temp->value = "";
            temp = temp->right;       
        }
        //FOR UPDATING THE GRID
        for (int i = 0; i < MAX_COLS; ++i)
        {
            grid[currentRow][i] = "";
        }
        
    }
    void clearColumn()
    {
        NodeCell<T>* temp = current;
        while (temp->up != nullptr)
        {
            temp = temp->up;
        }
        while (temp != nullptr)
        {
            temp->value = "";
            temp = temp->down;
        }
        //FOR UPDATING THE GRID
        for (int i = 0; i < MAX_ROWS; ++i)
        {
            grid[i][currentCol] = "";
        }
    }


    void deleteColumn()
    {
        if (MAX_COLS <= 1)
        {
            return;
        }

        NodeCell<T>* temp = current;       
        while (temp->up != nullptr)
        {
            temp = temp->up;
        }
        NodeCell<T>* delete_cell;

        if (temp->left == nullptr)
        {
            current = current->right;
            while (temp != nullptr)
            {
                delete_cell = temp;
                temp->right->left = nullptr;
                temp = temp->down;
                delete delete_cell;
            }
            for (int i = 0; i < MAX_ROWS; ++i)
            {
                grid[i].erase(grid[i].begin());
            }
        }

        else if (temp->right == nullptr)
        {
            currentCol--;
            current = current->left;
            while (temp != nullptr)
            {
                delete_cell = temp;
                temp->left->right = nullptr;
                temp = temp->down;
                delete delete_cell;
            }
            for (int i = 0; i < MAX_ROWS; ++i)
            {
                grid[i].pop_back();
            }
         
        }

        else
        {
            for (int i = 0; i < MAX_ROWS; ++i)
            {
                if (currentCol < grid[i].size())
                {
                    grid[i].erase(grid[i].begin() + currentCol);
                }
            }
            currentCol--;
            current = current->left;
            while (temp != nullptr)
            {
                delete_cell = temp;
                temp->left->right = temp->right;
                temp->right->left = temp->left;
                temp = temp->down;
                delete delete_cell;
            }
           
        }
               
        MAX_COLS--;
              
    }

    void deleteRow()
    {
        if (MAX_ROWS <= 1)
        {
            return;
        }
        NodeCell<T>* temp = current;
        while (temp->left != nullptr)
        {
            temp = temp->left;
        }
        NodeCell<T>* deleteCell;
        if (temp->up == nullptr)
        {
            current = current->down;
            while (temp != nullptr)
            {
                deleteCell = temp;
                temp->down->up = nullptr;
                temp = temp->right;
                delete deleteCell;
            }
            grid.erase(grid.begin());
        }

        else if (temp->down == nullptr)
        {
            currentRow--;
            current = current->up;
            while (temp != nullptr)
            {
                deleteCell = temp;
                temp->up->down = nullptr;
                temp = temp->right;
                delete deleteCell;
            }
            grid.pop_back();
         
        }

        else
        {
            currentRow--;
            current = current->up;
            while (temp != nullptr)
            {
                deleteCell = temp;
                temp->up->down = temp->down;
                temp->down->up = temp->up;
                temp = temp->right;
                delete deleteCell;
            }
            grid.erase(grid.begin() + currentRow+1);
         
        }
        MAX_ROWS--;
    }
  

    void printDebug()
    {
        cout << "MAX_ROWS: " << MAX_ROWS << endl;
        cout << "MAX_COLS: " << MAX_COLS << endl;
        cout << "current->value: " << current->value << endl;
        cout << "current->up: " << current->up << endl;
        cout << "current->down: " << current->down << endl;
        cout << "current->left: " << current->left << endl;
        cout << "current->right: " << current->right << endl;
    }

    
    NodeCell<T>* getNodeAt(T row, T col)
    {
        int rows = stoi(row);
        int cols = stoi(col);
        NodeCell<T>* temp = current;
        while (temp->left != nullptr)
        {
            temp = temp->left;
        }
        while (temp->up != nullptr)
        {
            temp = temp->up;
        }
        for (int i = 1; i < rows; ++i)
        {
            if (temp->down != nullptr)
            {
                temp = temp->down;
            }
        }
        for (int j = 1; j < cols; ++j)
        {
            if (temp->right != nullptr)
            {
                temp = temp->right;
            }
        }

        return temp;
    }

     T sumTotal(NodeCell<T>* startcell, NodeCell<T>* endcell , T srow ,T scol , T erow , T ecol)
     {
        int sum = 0;
        int summation = 0;

        if (startcell == nullptr || endcell == nullptr) {
            cout << "Invalid start or end cell." << endl;
        }

        NodeCell<T>* current = startcell;
        if (srow == erow)
        {
            while (current != nullptr && current != endcell->right)
            {              
                try
                {
                    summation = stoi(current->value);
                    sum = sum + summation;
                }
                catch (const std::invalid_argument& e)
                {
                    // Handle the case where stoi fails (e.g., empty string)
                }

                current = current->right;
            }
        }
        else if (ecol == scol)
        {
            while (current != nullptr && current != endcell->down)
            {
                try
                {
                    summation = stoi(current->value);
                    sum = sum + summation;
                }
                catch (const std::invalid_argument& e)
                {
                    // Handle the case where stoi fails (e.g., empty string)
                }

                current = current->down;
            }
        }
        else 
        {
            while (current != nullptr && current != endcell->right && current != endcell->down)
            {
                try
                {
                    summation = stoi(current->value);
                    sum += summation;
                }
                catch (const std::invalid_argument& e)
                {
                   // Handle the case where stoi fails (e.g., empty string)
                }
                current = current->right;

                if (current == nullptr || current == endcell->right)
                {
                    current = startcell->down;
                    startcell = startcell->down;
                }
            }
        }

        cout << "Total sum: " << sum << endl;

        return to_string(sum);
     }

     T calculateRangeSum(NodeCell<T>* startcell, NodeCell<T>* endcell)
     {
         int sum = 0;
         int summation = 0;
         NodeCell<T>* current = startcell;
         while (current != nullptr && current != endcell->right && current != endcell->down)
         {
             try
             {
                 summation = stoi(current->value);
                 sum += summation;
             }
             catch (const std::invalid_argument& e)
             {
                 // Handle the case where stoi fails (e.g., empty string)
             }
             current = current->right;

             if (current == nullptr || current == endcell->right)
             {
                 current = startcell->down;
                 startcell = startcell->down;
             }
         }

         cout << "Range sum: " << sum << endl;

         return to_string(sum);
     }

     T calculateAverage(NodeCell<T>* startcell, NodeCell<T>* endcell, T srow, T scol, T erow, T ecol)
     {
         NodeCell<T>* current = startcell;
         int count = 0;
         string sums = sumTotal(startcell, endcell, srow, scol, erow, ecol);
         int sum = stoi(sums);
         int avg = 0;
         int summation = 0;
        
         if (srow == erow)
         {
             while (current != nullptr && current != endcell->right)
             {
                 try
                 {
                     summation = stoi(current->value);
                     count++;
                 }
                 catch (const std::invalid_argument& e)
                 {
                     // Handle the case where stoi fails (e.g., empty string)
                 }

                 current = current->right;
             }
         }
         else if (ecol == scol)
         {
             while (current != nullptr && current != endcell->down)
             {
                 try
                 {
                     summation = stoi(current->value);
                     count++;
                 }
                 catch (const std::invalid_argument& e)
                 {
                     // Handle the case where stoi fails (e.g., empty string)
                 }

                 current = current->down;
             }
         }
         else
         {
             while (current != nullptr && current != endcell->right && current != endcell->down)
             {
                 try
                 {
                     summation = stoi(current->value);
                     count++;
                 }
                 catch (const std::invalid_argument& e)
                 {
                     // Handle the case where stoi fails (e.g., empty string)
                 }
                 current = current->right;

                 if (current == nullptr || current == endcell->right)
                 {
                     current = startcell->down;
                     startcell = startcell->down;
                 }
             }
         }
         avg = sum / count;
       
         cout << "Average " << avg << endl;

         return to_string(avg);

     }

     T calculateRangeAverage(NodeCell<T>* startcell, NodeCell<T>* endcell)
     {
         NodeCell<T>* current = startcell;
         int count = 0;
         int sum = 0;
         int avg = 0;
         int summation = 0;
         while (current != nullptr && current != endcell->right && current != endcell->down)
         {
             try
             {
                 summation = stoi(current->value);
                 sum  = sum +  summation;
                 count++;
             }
             catch (const std::invalid_argument& e)
             {
                 // Handle the case where stoi fails (e.g., empty string)
             }
             current = current->right;

             if (current == nullptr || current == endcell->right)
             {
                 current = startcell->down;
                 startcell = startcell->down;
             }
         }
         avg = sum / count;

         cout << "Range Average " << avg << endl;

         return to_string(avg);
     }

     T countNumbers(NodeCell<T>* startcell, NodeCell<T>* endcell, T srow, T scol, T erow, T ecol)
     {
         NodeCell<T>* current = startcell;
         int count = 0;
         int summation = 0;

         if (srow == erow)
         {
             while (current != nullptr && current != endcell->right)
             {
                 try
                 {
                     summation = stoi(current->value);
                     count++;
                 }
                 catch (const std::invalid_argument& e)
                 {
                     // Handle the case where stoi fails (e.g., empty string)
                 }

                 current = current->right;
             }
         }
         else if (ecol == scol)
         {
             while (current != nullptr && current != endcell->down)
             {
                 try
                 {
                     summation = stoi(current->value);
                     count++;
                 }
                 catch (const std::invalid_argument& e)
                 {
                     // Handle the case where stoi fails (e.g., empty string)
                 }

                 current = current->down;
             }
         }
         else
         {
             while (current != nullptr && current != endcell->right && current != endcell->down)
             {
                 try
                 {
                     summation = stoi(current->value);
                     count++;
                 }
                 catch (const std::invalid_argument& e)
                 {
                     // Handle the case where stoi fails (e.g., empty string)
                 }
                 current = current->right;

                 if (current == nullptr || current == endcell->right)
                 {
                     current = startcell->down;
                     startcell = startcell->down;
                 }
             }
         }

         cout << "Count : " << count << endl;

         return to_string(count);
        
     }

     T CalculateMax(NodeCell<T>* startcell, NodeCell<T>* endcell, T srow, T scol, T erow, T ecol)
     {
         NodeCell<T>* current = startcell;
         int max = -100000;
         int currentVal = 0;

         if (srow == erow)
         {
             while (current != nullptr && current != endcell->right)
             {
                 try
                 {
                     currentVal = stoi(current->value);
                     if (currentVal > max)
                     {
                         max = currentVal;
                     }
                 }
                 catch (const std::invalid_argument& e)
                 {
                     // Handle the case where stoi fails (e.g., empty string)
                 }

                 current = current->right;
                 
             }
         }
         else if (ecol == scol)
         {
             while (current != nullptr && current != endcell->down)
             {
                 try
                 {
                     currentVal = stoi(current->value);
                     if (currentVal > max)
                     {
                         max = currentVal;
                     }
                 }
                 catch (const std::invalid_argument& e)
                 {
                     // Handle the case where stoi fails (e.g., empty string)
                 }

                 current = current->down;
             }
         }
         else
         {
             while (current != nullptr && current != endcell->right && current != endcell->down)
             {
                 try
                 {
                     currentVal= stoi(current->value);
                     if (currentVal > max)
                     {
                         max = currentVal;
                     }
                 }
                 catch (const std::invalid_argument& e)
                 {
                     // Handle the case where stoi fails (e.g., empty string)
                 }
                 current = current->right;

                 if (current == nullptr || current == endcell->right)
                 {
                     current = startcell->down;
                     startcell = startcell->down;
                 }
             }
         }

         cout << "Max Value : " << max << endl;

         return to_string(max);
     
     }
     T calculateMin(NodeCell<T>* startcell, NodeCell<T>* endcell, T srow, T scol, T erow, T ecol)
     {
         NodeCell<T>* current = startcell;
         int min = 10000000;
         int currentVal = 0;

         if (srow == erow)
         {
             while (current != nullptr && current != endcell->right)
             {
                 try
                 {
                     currentVal = stoi(current->value);
                     if (currentVal < min)
                     {
                         min = currentVal;
                     }
                 }
                 catch (const std::invalid_argument& e)
                 {
                     // Handle the case where stoi fails (e.g., empty string)
                 }

                 current = current->right;
             }
         }
         else if (ecol == scol)
         {
             while (current != nullptr && current != endcell->down)
             {
                 try
                 {
                     currentVal = stoi(current->value);
                     if (currentVal < min)
                     {
                         min = currentVal;
                     }
                 }
                 catch (const std::invalid_argument& e)
                 {
                     // Handle the case where stoi fails (e.g., empty string)
                 }

                 current = current->down;
             }
         }
         else
         {
             while (current != nullptr && current != endcell->right && current != endcell->down)
             {
                 try
                 {
                     currentVal = stoi(current->value);
                     if (currentVal < min)
                     {
                         min = currentVal;
                     }
                 }
                 catch (const std::invalid_argument& e)
                 {
                     // Handle the case where stoi fails (e.g., empty string)
                 }
                 current = current->right;

                 if (current == nullptr || current == endcell->right)
                 {
                     current = startcell->down;
                     startcell = startcell->down;
                 }
             }
         }

         cout << "Min Value : " << min << endl;

         return to_string(min);
     }

     void copy(NodeCell<T>* startcell, NodeCell<T>* endcell, T srow, T scol, T erow, T ecol)
     {
         NodeCell<T>* current = startcell;
         clipboardVector.clear();
         if (srow == erow)
         {
             while (current != nullptr && current != endcell->right)
             {
                 
                 clipboardVector.push_back(current->value);
                 current = current->right;
             }
             cout << "Copied Successfully !!";
             _getch();
         }
         else if (ecol == scol)
         {
             while (current != nullptr && current != endcell->down)
             {
                 clipboardVector.push_back(current->value);
                 current = current->down;
             }
             cout << "Copied Successfully !!";
             _getch();
         }
         else
         {
             while (current != nullptr && current != endcell->right && current != endcell->down)
             {
                 clipboardVector.push_back(current->value);   
                 current = current->right;

                 if (current == nullptr || current == endcell->right)
                 {
                     current = startcell->down;
                     startcell = startcell->down;
                 }
             }
             cout << "Copied Successfully !!";
             _getch();
         }
         
     }

     void cut(char option)
     {
         NodeCell<T>* temp = current;
         clipboardVector.clear();
         if (option =='r' || option == 'R')
         {
             while (temp->left != nullptr)
             {
                 temp = temp->left;
             }
             while (temp != nullptr)
             {
                 clipboardVector.push_back(temp->value);
                 temp = temp->right;
             }
             clearRow();
             cout << "Cut Successfully !!";
             _getch();
         }
         else if (option == 'c' || option == 'C')
         {
             while (temp->up != nullptr)
             {
                 temp = temp->up;
             }
             while (temp != nullptr)
             {
                 clipboardVector.push_back(temp->value);
                 temp = temp->down;
             }
             clearColumn();
             cout << "Cut Successfully !!";
             _getch();
         }
         
     }
     NodeCell<T>* getNodeAtPaste(T row, T col)
     {
         int rows = stoi(row);
         int cols = stoi(col);
         NodeCell<T>* temp = current;
         while (temp->left != nullptr)
         {
             temp = temp->left;
         }
         while (temp->up != nullptr)
         {
             temp = temp->up;
         }
         for (int i = 0; i < rows; ++i)
         {
             if (temp->down != nullptr)
             {
                 temp = temp->down;
             }
         }
         for (int j = 0; j < cols; ++j)
         {
             if (temp->right != nullptr)
             {
                 temp = temp->right;
             }
         }

         return temp;
     }

     void paste(T srow, T scol, T erow, T ecol, T row, T col)
     {
         NodeCell<T>* finalNode = getNodeAt(row, col);
         NodeCell<T>* temp = current;
         current = finalNode;
         int rrow = stoi(row);
         int ccol = stoi(col);
         if (srow == erow)
         {
             while (!clipboardVector.empty())
             {
                 finalNode->value = clipboardVector.front();
                 clipboardVector.erase(clipboardVector.begin());
                 if (clipboardVector.empty())
                 {
                     break;
                 }
                 else if (finalNode->right != nullptr)
                 {
                     finalNode = finalNode->right;
                     current = finalNode;

                 }
                 else
                 {
                     insertCoulmnToRight();
                     finalNode = finalNode->right;
                     current = finalNode;

                 }

             }
             current = getNodeAt("0", "0");
             for (int i = 0; i < MAX_ROWS ; i++)
             {
                 for (int j = 0; j < MAX_COLS ; j++)
                 {
                    
                     grid[i][j] = getNodeAtPaste(to_string(i), to_string(j))->value;
                 }
             }
             current = temp;

            
             
         }
         else if (scol == ecol)
         {
             
             while (!clipboardVector.empty())
             {
                 finalNode->value = clipboardVector.front();
                 clipboardVector.erase(clipboardVector.begin());
                 if (clipboardVector.empty())
                 {
                     break;
                 }
                 if (finalNode->down != nullptr)
                 {
                     finalNode = finalNode->down;
                     current = finalNode;
                 }
                 else
                 {
                     InsertRowBelow();
                     finalNode = finalNode->down;
                     current = finalNode;

                 }

             }
             current = getNodeAt("0", "0");
             for (int i = 0; i < MAX_ROWS; i++)
             {
                 for (int j = 0; j < MAX_COLS; j++)
                 {
                     grid[i][j] = getNodeAtPaste(to_string(i), to_string(j))->value;
                 }
             }
             current = temp;
         }


     }

     
    void writeCalculation(string sum)
    {
        string row, col;
        NodeCell<T>* finalcell = new NodeCell<T>();
        cout << "Write It in Cell(Row,Col) : \n";
        cin >> row >> col;
        int rows = stoi(row) - 1;
        int cols = stoi(col) - 1;
        if (rows >= 0 && rows < MAX_ROWS && cols >= 0 && cols < MAX_COLS)
        {
            finalcell = getNodeAt(row, col);
            finalcell->value = sum;
            grid[rows][cols] = sum;
        }
        else
        {
            cout << "Invalid cell coordinates. No cell in the grid." << endl;
            _getch();
        }

    }

    void saveSheetToFile(const std::string& filename) {
        std::ofstream file(filename);

        if (!file.is_open()) {
            std::cerr << "Error: Could not open file for writing.\n";
            return;
        }

        for (int i = 0; i < MAX_ROWS; ++i) {
            for (int j = 0; j < MAX_COLS; ++j) {
                file << grid[i][j] << '\t'; // Assuming tab-separated values, adjust as needed
            }
            file << '\n';
        }

        file.close();
        cout << "Sheet Saved Successfully in mysheet.txt";
    }

    void loadSheetFromFile(const std::string& filename) {
        std::ifstream file(filename);

        if (!file.is_open()) {
            std::cerr << "Error: Could not open file for reading.\n";
            return;
        }

        // Clear the existing grid
        grid.clear();

        std::string line;
        while (std::getline(file, line)) {
            std::istringstream iss(line);
            std::vector<std::string> row;

            std::string value;
            while (std::getline(iss, value, '\t')) {
                row.push_back(value);
            }

            grid.push_back(row);
        }

        file.close();
    }

    void display()
    {
        
        std::cout << " .----------------.  .----------------.  .-----------------.  .----------------." << std::endl;
        std::cout << "| .--------------. || .--------------. || .--------------. || .--------------. |" << std::endl;
        std::cout << "| | ____    ____ | || |     _____    | || | ____  _____  | || |     _____    | |" << std::endl;
        std::cout << "| ||_   \\  /   _|| || |    |_   _|   | || ||_   \\|_   _| | || |    |_   _|   | |" << std::endl;
        std::cout << "| |  |   \\/   |  | || |      | |     | || |  |   \\ | |   | || |      | |     | |" << std::endl;
        std::cout << "| |  | |\\  /| |  | || |      | |     | || |  | |\\ \\| |   | || |      | |     | |" << std::endl;
        std::cout << "| | _| |_\\/_| |_ | || |     _| |_    | || | _| |_\\   |_  | || |     _| |_    | |" << std::endl;
        std::cout << "| ||_____||_____|| || |    |_____|   | || ||_____|\____| | || |    |_____|   | |" << std::endl;
        std::cout << "| |              | || |              | || |              | || |              | |" << std::endl;
        std::cout << "| '--------------' || '--------------' || '--------------' || '--------------' |" << std::endl;
        std::cout << " '----------------'  '----------------'  '----------------'  '----------------'" << std::endl;

        std::cout << std::endl; 

        std::cout << " .----------------.  .----------------.  .----------------.  .----------------.  .----------------." << std::endl;
        std::cout << "| .--------------. || .--------------. || .--------------. || .--------------. || .--------------. |" << std::endl;
        std::cout << "| |  _________   | || |  ____  ____  | || |     ______   | || |  _________   | || |   _____      | |" << std::endl;
        std::cout << "| | |_   ___  |  | || | |_  _||_  _| | || |   .' ___  |  | || | |_   ___  |  | || |  |_   _|     | |" << std::endl;
        std::cout << "| |   | |_  \\_|  | || |   \\ \\  / /   | || |  / .'   \\_|  | || |   | |_  \\_|  | || |    | |       | |" << std::endl;
        std::cout << "| |   |  _|  _   | || |    > `' <    | || |  | |         | || |   |  _|  _   | || |    | |   _   | |" << std::endl;
        std::cout << "| |  _| |___/ |  | || |  _/ /'`\\ \\_  | || |  \\ `.___.'\\  | || |  _| |___/ |  | || |   _| |__/ |  | |" << std::endl;
        std::cout << "| | |_________|  | || | |____||____| | || |   `._____.'  | || | |_________|  | || |  |________|  | |" << std::endl;
        std::cout << "| |              | || |              | || |              | || |              | || |              | |" << std::endl;
        std::cout << "| '--------------' || '--------------' || '--------------' || '--------------' || '--------------' |" << std::endl;
        std::cout << " '----------------'  '----------------'  '----------------'  '----------------'  '----------------'" << std::endl;
    }

    
};


int main()
{
    
    MiniExcelClass <string> excelSheet;
    excelSheet.display();
    char cutoption;
    const int totalSteps = 100;
    const int barWidth = 50;

    std::cout << "Loading...\n";

    for (int i = 0; i <= totalSteps; ++i) {
        printLoadingBar(i, totalSteps, barWidth);
        std::this_thread::sleep_for(std::chrono::milliseconds(50));  // Simulate some work being done
    }

    std::cout << "\nLoading complete! Project starting...\n";
    _getch();
    system("cls");
    SetConsoleTextAttribute(acolor, FOREGROUND_RED | FOREGROUND_GREEN | FOREGROUND_BLUE);
    MiniExcelClass<string>::NodeCell<string>* startcell = new MiniExcelClass<string>::NodeCell<string>();
    MiniExcelClass<string>::NodeCell<string>* endcell = new MiniExcelClass<string>::NodeCell<string>();
    string startrow, startcol, endrow, endcol , sum , average , countTotal , maxNum , minNum , rangeavg , rangesum;
    string clipStartRow, clipStartCol, clipEndRow, clipEndCol , rowclip , columnclip;
    

    // Load the sheet from the file
    // excelSheet.loadSheetFromFile("mysheet.txt");

    while (true) {
        excelSheet.displaySheet();
        excelSheet.getCurrentNode();
        cout << endl;
        std::cout << std::left << std::setw(15) << "Command" << std::setw(35) << "Description" << std::setw(15) << "Command" << std::setw(35) << "Description" << std::endl;
        std::cout << std::string(105, '-') << std::endl;

        std::cout << std::left << std::setw(15) << "W" << std::setw(35) << "Write" << std::setw(15) << "U" << std::setw(35) << "Move Up" << std::endl;
        std::cout << std::left << std::setw(15) << "D" << std::setw(35) << "Move Down" << std::setw(15) << "L" << std::setw(35) << "Move Left" << std::endl;
        std::cout << std::left << std::setw(15) << "R" << std::setw(35) << "Move Right" << std::setw(15) << "A" << std::setw(35) << "Add Row Above" << std::endl;
        std::cout << std::left << std::setw(15) << "B" << std::setw(35) << "Add Row Below" << std::setw(15) << "C" << std::setw(35) << "Add Columns Right" << std::endl;
        std::cout << std::left << std::setw(15) << "E" << std::setw(35) << "Add Column Left" << std::setw(15) << "S" << std::setw(35) << "Insert Cell by Right Shift" << std::endl;
        std::cout << std::left << std::setw(15) << "K" << std::setw(35) << "Insert Cell by Down Shift" << std::setw(15) << "M" << std::setw(35) << "Delete Cell by Left Shift" << std::endl;
        std::cout << std::left << std::setw(15) << "N" << std::setw(35) << "Delete Cell by Up Shift" << std::setw(15) << "X" << std::setw(35) << "Delete Row" << std::endl;
        std::cout << std::left << std::setw(15) << "Y" << std::setw(35) << "Delete Column" << std::setw(15) << "G" << std::setw(35) << "Clear Row" << std::endl;
        std::cout << std::left << std::setw(15) << "H" << std::setw(35) << "Clear Column" << std::setw(15) << "O" << std::setw(35) << "Calculate Sum" << std::endl;
        std::cout << std::left << std::setw(15) << "V" << std::setw(35) << "Calculate Average" << std::setw(15) << "T" << std::setw(35) << "Count Numbers" << std::endl;
        std::cout << std::left << std::setw(15) << "P" << std::setw(35) << "Max Number" << std::setw(15) << "Q" << std::setw(35) << "Min Number" << std::endl;
        std::cout << std::left << std::setw(15) << "F" << std::setw(35) << "Range Sum" << std::setw(15) << "I" << std::setw(35) << "Range Average" << std::endl;
        std::cout << std::left << std::setw(15) << "+" << std::setw(35) << "Copy" << std::setw(15) << "-" << std::setw(35) << "Cut" << std::endl;
        std::cout << std::left << std::setw(15) << "/" << std::setw(35) << "Paste" << std::setw(15) << "." << std::setw(35) << "Save Sheet" << std::endl;
        std::cout << std::left << std::setw(15) << "Z" << std::setw(35) << "Quit" << std::setw(15) << "" << std::setw(35) << "" << std::endl;
        char command;
        cin >> command;

        switch (command) {
        case 'W':
        case 'w': {
            std::string data;
            std::cout << "Enter data: ";
            std::cin >> data;
            excelSheet.current->value = data;
            excelSheet.writeData(data);
            break;
        }
        case 'U':
        case 'u':
            excelSheet.moveUp();
            excelSheet.printDebug();            
            _getch();
            break;
        case 'D':
        case 'd':
            excelSheet.moveDown();
            excelSheet.printDebug();
            _getch();
            break;
        case 'L':
        case 'l':
            excelSheet.moveLeft();
            excelSheet.printDebug();
            _getch();
            break;
        case 'R':
        case 'r':
            excelSheet.moveRight();
            excelSheet.printDebug();
            _getch();
            break;
        case 'B':
        case 'b':
           excelSheet.InsertRowBelow();
            break;
        case 'C':
        case 'c':
            excelSheet.insertCoulmnToRight();
            break;
        case 'A':
        case 'a':
            excelSheet.insertRowAbove();
            break;
        case 'E':
        case 'e':
            excelSheet.insertColumntoLeft();
            break;
        case 'S':
        case 's':
            excelSheet.insertCellByRightShift();
            break;
        case 'K':
        case 'k':
            excelSheet.insertCellByDownShift();
            break;
        case 'M':
        case 'm':
            excelSheet.deleteCellbyLeftShift();
            break;
        case 'N':
        case 'n':
            excelSheet.deleteCellbyUpShift();
            break;
        case 'O':
        case 'o':
            cout << "Enter starting row & column : \n";
            cin >> startrow >> startcol;
            cout << "Enter ending row & column : \n";
            cin >> endrow >> endcol;
            startcell =  excelSheet.getNodeAt(startrow , startcol);
            endcell = excelSheet.getNodeAt(endrow, endcol);
            sum = excelSheet.sumTotal(startcell, endcell , startrow, startcol , endrow , endcol);
            excelSheet.writeCalculation(sum);
            break;
        case 'Y':
        case 'y':
            excelSheet.deleteColumn();
            break;
        case 'X':
        case 'x':
            excelSheet.deleteRow();
            break;
        case 'G':
        case 'g':
            excelSheet.clearRow();
            break;
        case 'H':
        case 'h':
            excelSheet.clearColumn();
            break;
        case 'V':
        case 'v':
            cout << "Enter starting row & column : \n";
            cin >> startrow >> startcol;
            cout << "Enter ending row & column : \n";
            cin >> endrow >> endcol;
            startcell = excelSheet.getNodeAt(startrow, startcol);
            endcell = excelSheet.getNodeAt(endrow, endcol);
            average = excelSheet.calculateAverage(startcell, endcell, startrow, startcol, endrow, endcol);
            excelSheet.writeCalculation(average);
            break;
        case 'T':
        case 't':
            cout << "Enter starting row & column : \n";
            cin >> startrow >> startcol;
            cout << "Enter ending row & column : \n";
            cin >> endrow >> endcol;
            startcell = excelSheet.getNodeAt(startrow, startcol);
            endcell = excelSheet.getNodeAt(endrow, endcol);
            countTotal = excelSheet.countNumbers(startcell, endcell, startrow, startcol, endrow, endcol);
            excelSheet.writeCalculation(countTotal);
            break;
        case 'P':
        case 'p':
            cout << "Enter starting row & column : \n";
            cin >> startrow >> startcol;
            cout << "Enter ending row & column : \n";
            cin >> endrow >> endcol;
            startcell = excelSheet.getNodeAt(startrow, startcol);
            endcell = excelSheet.getNodeAt(endrow, endcol);
            maxNum = excelSheet.CalculateMax(startcell, endcell, startrow, startcol, endrow, endcol);
            excelSheet.writeCalculation(maxNum);
            break;
        case 'Q':
        case 'q':
            cout << "Enter starting row & column : \n";
            cin >> startrow >> startcol;
            cout << "Enter ending row & column : \n";
            cin >> endrow >> endcol;
            startcell = excelSheet.getNodeAt(startrow, startcol);
            endcell = excelSheet.getNodeAt(endrow, endcol);
            minNum= excelSheet.calculateMin(startcell, endcell, startrow, startcol, endrow, endcol);
            excelSheet.writeCalculation(minNum);
            break;
        case 'I':
        case 'i':
            cout << "Enter starting row & column : \n";
            cin >> startrow >> startcol;
            cout << "Enter ending row & column : \n";
            cin >> endrow >> endcol;
            startcell = excelSheet.getNodeAt(startrow, startcol);
            endcell = excelSheet.getNodeAt(endrow, endcol);
            rangeavg= excelSheet.calculateRangeAverage(startcell, endcell);
            excelSheet.writeCalculation(rangeavg);
            break;
        case 'F':
        case 'f':
            cout << "Enter starting row & column : \n";
            cin >> startrow >> startcol;
            cout << "Enter ending row & column : \n";
            cin >> endrow >> endcol;
            startcell = excelSheet.getNodeAt(startrow, startcol);
            endcell = excelSheet.getNodeAt(endrow, endcol);
            rangesum = excelSheet.calculateRangeSum(startcell, endcell);
            excelSheet.writeCalculation(rangesum);           
            break;
        case '+':
            cout << "Enter starting row & column : \n";
            cin >> clipStartRow>> clipStartCol;
            cout << "Enter ending row & column : \n";
            cin >> clipEndRow >>clipEndCol;
            startcell = excelSheet.getNodeAt(clipStartRow, clipStartCol);
            endcell = excelSheet.getNodeAt(clipEndRow, clipEndCol);
            excelSheet.copy(startcell, endcell, clipStartRow, clipStartCol, clipEndRow, clipEndCol);
            break;
        case '-':
            cout << "Want to cut the Entire Row or Column (r/c) : ";
            cin >> cutoption;
            excelSheet.cut(cutoption);
            break;
        case '/':
            cout << "Enter from which starting cell you want to Paste : \n";
            cout << "Row : ";
            cin >> rowclip;
            cout << "Column : ";
            cin >> columnclip;
            excelSheet.paste(clipStartRow,clipStartCol , clipEndRow , clipEndCol , rowclip , columnclip);
            break;
        case '.':
            excelSheet.saveSheetToFile("mysheet.txt");
            _getch();
            break;
        case 'Z':
        case 'z':
            return 0;
        default:
            std::cout << "Invalid command. Try again." << std::endl;
        }
       
    }

    return 0;
}



