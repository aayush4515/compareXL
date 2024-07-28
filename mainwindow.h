#ifndef MAINWINDOW_H
#define MAINWINDOW_H

#include <QMainWindow>
#include <QComboBox>
#include <QPushButton>
#include <QTextEdit>
#include <string>

// LibXL library
#include "/Users/aayush/Desktop/QtProjects/compareXL/include_cpp/enum.h"
#include "/Users/aayush/Desktop/QtProjects/compareXL/include_cpp/IAutoFilterT.h"
#include "/Users/aayush/Desktop/QtProjects/compareXL/include_cpp/IBookT.h"
#include "/Users/aayush/Desktop/QtProjects/compareXL/include_cpp/IConditionalFormatT.h"
#include "/Users/aayush/Desktop/QtProjects/compareXL/include_cpp/IConditionalFormattingT.h"
#include "/Users/aayush/Desktop/QtProjects/compareXL/include_cpp/IFilterColumnT.h"
#include "/Users/aayush/Desktop/QtProjects/compareXL/include_cpp/IFontT.h"
#include "/Users/aayush/Desktop/QtProjects/compareXL/include_cpp/IFormatT.h"
#include "/Users/aayush/Desktop/QtProjects/compareXL/include_cpp/IFormControlT.h"
#include "/Users/aayush/Desktop/QtProjects/compareXL/include_cpp/IRichStringT.h"
#include "/Users/aayush/Desktop/QtProjects/compareXL/include_cpp/ISheetT.h"
#include "/Users/aayush/Desktop/QtProjects/compareXL/include_cpp/libxl.h"
#include "/Users/aayush/Desktop/QtProjects/compareXL/include_cpp/setup.h"

using namespace libxl;
using namespace std;

QT_BEGIN_NAMESPACE
namespace Ui {
class MainWindow;
}
QT_END_NAMESPACE

class MainWindow : public QMainWindow
{
    Q_OBJECT

public:
    MainWindow(QWidget *parent = nullptr);
    ~MainWindow();

private slots:
    void onCompareClicked();
    void loadOldFile();
    void loadNewFile();

    // void onOldRowSelectionChanged(int index);
    // void onOldColSelectionChanged(int index);
    // void onNewRowSelectionChanged(int index);
    // void onNewColSelectionChanged(int index);

private:
    Ui::MainWindow *ui;
    QComboBox *comboBoxOldSheet;
    QComboBox *comboBoxNewSheet;

    QComboBox *oldRowComboBox;
    QComboBox *oldColComboBox;
    QComboBox *newRowComboBox;
    QComboBox *newColComboBox;

    QPushButton *compareButton;
    QTextEdit *resultTextEdit;
    Book* oldBook;
    Book* newBook;

    // variables to store old and new row and col positions
    int oldRow;
    int oldCol;
    int newRow;
    int newCol;

    // containers to store student IDs
    vector<string> oldIDs;
    vector<string> newIDs;
    vector<string> addOns;


    //void populateComboBoxes(Book* oldBook, Book* newBook);
    void populateComboBoxes(QComboBox* comboBox, Book* book);
    void populateStartingComboBoxes();
    void compareSheets();
    int getSheetIndexByName(Book* book, const string& sheetName);

    // method to clear previous data
    void clearPreviousData();

    // methods to set old and new row and col positions
    void setOldRow();
    void setOldCol();
    void setNewRow();
    void setNewCol();


};
#endif // MAINWINDOW_H
