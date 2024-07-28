#include "mainwindow.h"
#include "./ui_mainwindow.h"

#include <QFileDialog>
#include <QMessageBox>
#include <algorithm>
#include <QDebug>
#include <string>

using namespace std;

MainWindow::MainWindow(QWidget *parent)
    : QMainWindow(parent)
    , ui(new Ui::MainWindow)
    , oldBook(nullptr)
    , newBook(nullptr)
{
    ui->setupUi(this);

    comboBoxOldSheet = ui->comboBox_1;
    comboBoxNewSheet = ui->comboBox_2;
    oldRowComboBox = ui->oldRowComboBox;
    oldColComboBox = ui->oldColComboBox;
    newRowComboBox = ui->newRowComboBox;
    newColComboBox = ui->newColComboBox;

    resultTextEdit = ui->resultTextEdit;
    compareButton = ui->compareButton;


    // populate the starting row and column dropdown box as soon as the ui loads
    populateStartingComboBoxes();


    // connect(ui->oldRowComboBox, QOverload<int>::of(&QComboBox::currentIndexChanged), this, &MainWindow::onOldRowSelectionChanged);
    // connect(ui->oldColComboBox, QOverload<int>::of(&QComboBox::currentIndexChanged), this, &MainWindow::onOldColSelectionChanged);
    // connect(ui->newRowComboBox, QOverload<int>::of(&QComboBox::currentIndexChanged), this, &MainWindow::onNewRowSelectionChanged);
    // connect(ui->newColComboBox, QOverload<int>::of(&QComboBox::currentIndexChanged), this, &MainWindow::onNewColSelectionChanged);


    connect(ui->compareButton, &QPushButton::clicked, this, &MainWindow::onCompareClicked);
    connect(ui->loadOldFilesButton, &QPushButton::clicked, this, &MainWindow::loadOldFile);
    connect(ui->loadNewFilesButton, &QPushButton::clicked, this, &MainWindow::loadNewFile);

}

MainWindow::~MainWindow()
{
    // Ensure books are released when the application is closed
    if (this->oldBook) {
        this->oldBook->release();
    }
    if (this->newBook) {
        this->newBook->release();
    }

    delete ui;
}

void MainWindow::onCompareClicked() {

    // set the old and new row and col values after user selection
    setOldRow();
    setOldCol();
    setNewRow();
    setNewCol();

    compareSheets();
}

void MainWindow::clearPreviousData() {

    // clear the vectors to make it available for new data
    oldIDs.clear();
    newIDs.clear();
    addOns.clear();


    // // release the new book and old book
    // if (this->oldBook) {
    //     this->oldBook->release();
    //     this->oldBook = nullptr;
    // }
    // if (this->newBook) {
    //     this->newBook->release();
    //     this->newBook = nullptr;
    // }

}

void MainWindow::setOldRow() {
    bool ok;
    this->oldRow = ui->oldRowComboBox->currentText().toInt(&ok, 10);

    if (ok) {
        // Conversion succeeded, number contains the integer value
        qDebug() << "The integer value is:" << this->oldRow;;
    } else {
        // Conversion failed
        qDebug() << "Conversion failed";
    }
}

void MainWindow::setOldCol() {
    bool ok;
    this->oldCol = ui->oldColComboBox->currentText().toInt(&ok, 10);

    if (ok) {
        // Conversion succeeded, number contains the integer value
        qDebug() << "The integer value is:" << this->oldCol;
    } else {
        // Conversion failed
        qDebug() << "Conversion failed";
    }
}

void MainWindow::setNewRow() {
    bool ok;
    this->newRow = ui->newRowComboBox->currentText().toInt(&ok, 10);

    if (ok) {
        // Conversion succeeded, number contains the integer value
        qDebug() << "The integer value is:" << this->newRow;
    } else {
        // Conversion failed
        qDebug() << "Conversion failed";
    }
}

void MainWindow::setNewCol() {
    bool ok;
    this->newCol = ui->newColComboBox->currentText().toInt(&ok, 10);

    if (ok) {
        // Conversion succeeded, number contains the integer value
        qDebug() << "The integer value is:" << this->newCol;
    } else {
        // Conversion failed
        qDebug() << "Conversion failed";
    }
}


// void MainWindow::onOldRowSelectionChanged(int index) {
//     QString selectedText = oldRowComboBox->itemText(index);
//     ui->oldRowComboBox->setItemText(index, selectedText);           // added by me to try
//     qDebug() << "Selected Old Row:" << selectedText;
// }

// void MainWindow::onOldColSelectionChanged(int index) {
//     QString selectedText = oldColComboBox->itemText(index);
//     qDebug() << "Selected Old Column:" << selectedText;
// }

// void MainWindow::onNewRowSelectionChanged(int index) {
//     QString selectedText = newRowComboBox->itemText(index);
//     qDebug() << "Selected New Row:" << selectedText;
// }

// void MainWindow::onNewColSelectionChanged(int index) {
//     QString selectedText = newColComboBox->itemText(index);
//     qDebug() << "Selected New Column:" << selectedText;
// }

// populates the old and new starting rows and cols
void MainWindow::populateStartingComboBoxes() {

    for (int i = 1; i <= 10; i++) {
        oldRowComboBox->addItem(QString::number(i));    // converts i to a QString and adds to the combo box
        oldColComboBox->addItem(QString::number(i));
        newRowComboBox->addItem(QString::number(i));
        newColComboBox->addItem(QString::number(i));
    }

}


void MainWindow::loadOldFile() {
    // Use QFileDialog to select and load the old Excel file
    QString oldFilePath = QFileDialog::getOpenFileName(this, "Open Old Excel File", "", "Excel Files (*.xls *.xlsx)");

    if (oldFilePath.isEmpty()) {
        QMessageBox::warning(this, "Warning", "Failed to load the old Excel File.");
        return;
    }

    // Load the Excel file using libxl
    this->oldBook = xlCreateXMLBook();

    if (!this->oldBook->load(oldFilePath.toStdString().c_str())) {
        QMessageBox::warning(this, "Warning", "Failed to load the old Excel file.");
        return;
    }

    populateComboBoxes(comboBoxOldSheet, this->oldBook);
}

void MainWindow::loadNewFile() {
    // Use QFileDialog to select and load the new Excel file
    QString newFilePath = QFileDialog::getOpenFileName(this, "Open New Excel File", "", "Excel Files (*.xls *.xlsx)");

    if (newFilePath.isEmpty()) {
        QMessageBox::warning(this, "Warning", "Failed to load the new Excel File.");
        return;
    }

    // Load the Excel file using libxl
    this->newBook = xlCreateXMLBook();

    if (!this->newBook->load(newFilePath.toStdString().c_str())) {
        QMessageBox::warning(this, "Warning", "Failed to load the new Excel file.");
        return;
    }

    populateComboBoxes(comboBoxNewSheet, this->newBook);
}


void MainWindow::populateComboBoxes(QComboBox* comboBox, Book* book) {
    // Populate combo box with sheet names from the loaded Excel file
    comboBox->clear(); // Clear existing items

    if (!book) {
        QMessageBox::warning(this, "Warning", "Book object is null.");
        return;
    }


    for (int i = 0; i < book->sheetCount(); ++i) {
        Sheet* sheet = book->getSheet(i);
        if (sheet) {
            comboBox->addItem(sheet->name());
        }
    }
}

int MainWindow::getSheetIndexByName(Book* book, const string& sheetName) {
    if (!book) {
        QMessageBox::warning(this, "Warning", "Book object is null.");
        return -1;
    }


    for (int i = 0; i < book->sheetCount(); i++) {
        Sheet* currSheet = book->getSheet(i);
        if (currSheet && sheetName == currSheet->name()) {
            return i;
        }
    }
    return -1; // sheet not found
}


void MainWindow::compareSheets() {
    // Compare the selected sheets and display the results
    QString oldSheetName = comboBoxOldSheet->currentText();
    QString newSheetName = comboBoxNewSheet->currentText();


    // // Load selected sheets
    // Sheet* oldSheet = this->oldBook->getSheet(oldSheetName.toStdString().c_str());
    // Sheet* newSheet = this->newBook->getSheet(newSheetName.toStdString().c_str());

    // Find the index of the selected sheets
    int oldSheetIndex = getSheetIndexByName(this->oldBook, oldSheetName.toStdString());
    int newSheetIndex = getSheetIndexByName(this->newBook, newSheetName.toStdString());

    if (oldSheetIndex == -1 || newSheetIndex == -1) {
        QMessageBox::warning(this, "Warning", "Sheet not found in one of the books.");
        return;
    }


    // Load the selected sheets
    Sheet* oldSheet = this->oldBook->getSheet(oldSheetIndex);
    Sheet* newSheet = this->newBook->getSheet(newSheetIndex);


    if (!oldSheet || !newSheet) {
        QMessageBox::warning(this, "Warning", "Failed to load the selected sheets.");
        return;
    }


    // Extract the student IDs and compare


    for (int row = this->oldRow - 1; row < oldSheet->lastRow(); ++row) {
        const char* cellValue = oldSheet->readStr(row, this->oldCol - 1);
        if (cellValue) { // Check if cellValue is not null
            oldIDs.push_back(cellValue);
        }
    }

    for (int row = this->newRow - 1; row < newSheet->lastRow(); ++row) {
        const char* cellValue = newSheet->readStr(row, this->newCol - 1);
        if (cellValue) { // Check if cellValue is not null
            newIDs.push_back(cellValue);
        }
        // else {
        //     qDebug() << "Row" << row << " Cell value is null.";
        // }
    }


    // finding add-ons
    for (const auto& id : newIDs) {
        auto it = std::find(oldIDs.begin(), oldIDs.end(), id);
        if (it == oldIDs.end()) {
            addOns.push_back(id);
        }
    }

    //display results
    resultTextEdit->clear();
    for (const auto &id : addOns) {
        resultTextEdit->append(QString::fromStdString(id));
    }

    resultTextEdit->update();
    clearPreviousData();

    // // displaying newIDs list for testing:
    // resultTextEdit->clear();
    // for (const auto &id : newIDs) {
    //     resultTextEdit->append(QString::fromStdString(id));
    // }


}
