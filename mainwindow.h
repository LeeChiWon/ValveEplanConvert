#ifndef MAINWINDOW_H
#define MAINWINDOW_H

#include <QMainWindow>
#include <QtWidgets>
#include <QtXlsx>
#include "libxl.h"

#define LINECOLOR 0
#define IDENTITY 1
#define VALVELINENUMBER 2
#define VALVENUMBER 3

using namespace libxl;

namespace Ui {
class MainWindow;
}

class MainWindow : public QMainWindow
{
    Q_OBJECT

public:
    explicit MainWindow(QWidget *parent = 0);
    ~MainWindow();

    QTranslator Trans;
    QMultiMap<QString,QString> ValveBlock;
    QProcess *Process;

    void ValveFileRead(QString);
    void DrawingFileWrite(QString);

protected:
    void dragEnterEvent(QDragEnterEvent *event);
    void dropEvent(QDropEvent *event);

private slots:
    void Statusbar_Show(QString);
    void on_pushButton_Delete_clicked();
    void on_pushButton_Convert_clicked();
    void on_pushButton_DeleteAll_clicked();
    void on_actionHelp_triggered();
    void on_pushButton_Open_clicked();

private:
    Ui::MainWindow *ui;
};

#endif // MAINWINDOW_H
