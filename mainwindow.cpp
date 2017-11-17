#include "mainwindow.h"
#include "ui_mainwindow.h"

MainWindow::MainWindow(QWidget *parent) :
    QMainWindow(parent),
    ui(new Ui::MainWindow)
{
    ui->setupUi(this);

    Trans.load(":/Lang/Lang_ko_KR.qm");
    QApplication::installTranslator(&Trans);
    ui->retranslateUi(this);

    setAcceptDrops(true);

    // QXlsx::Document xlsx(tr("S:/ControlDATA/Document/Valve/Country/SMC/E16_272_miraeENG_6T_31.xls"));
    /*QXlsx::Document xlsx(tr("D:/E16_272_miraeENG_6T_31.xlsx"));
    qDebug()<<xlsx.read("B38");
    xlsx.deleteLater();*/

}

MainWindow::~MainWindow()
{
    delete ui;
}

void MainWindow::ValveFileRead(QString FileName)
{
    QFile File(FileName);
    QStringList LineText;
    int Count=0;
    QString FileLineText,tmp;

    if(File.open(QFile::ReadOnly|QFile::Text))
    {
        QTextStream OpenFile(&File);

        while(!OpenFile.atEnd())
        {
            FileLineText=OpenFile.readLine().replace(" ","");
            LineText=FileLineText.split(",");//,QString::SkipEmptyParts);

            if(LineText.count()<1)
            {
                continue;
            }

            if(LineText.at(0)=="NO.")
            {
                Count++;
                continue;
            }
            tmp=LineText.at(1);
            ValveBlock.insert(tmp.replace("-",""),QString("VB-%1").arg(Count));
            ValveBlock.insert(tmp.replace("-",""),LineText.at(0));
            ValveBlock.insert(tmp.replace("-",""),LineText.at(1));
            ValveBlock.insert(tmp.replace("-",""),LineText.at(2));

            //qDebug()<<LineText;
        }
        //qDebug()<<ValveBlock;
        File.close();
    }
}

void MainWindow::DrawingFileWrite(QString FileName)
{
    qDebug()<<FileName;
    Book *book=xlCreateBook();
    QByteArray byte=FileName.toLocal8Bit();
    char* File_Name=byte.data();
    char *Text;
    int RowCount=2;
    QList<QString> values;
    QStringList Split;
    QString Tmp;
    bool check;

    if(book->load(File_Name))
    {
        Sheet* sheet=book->getSheet(0);
        while(1)
        {
            Tmp=sheet->readStr(RowCount,4);
            if(Tmp=="" ||Tmp==" " ||Tmp==NULL)
            {
                break;
            }            
            if(ValveBlock.contains(Tmp))
            {
                values=ValveBlock.values(Tmp);
                qDebug()<<values;
                for(int i=0; i<values.size(); i++)
                {
                    switch(i)
                    {
                    case LINECOLOR:
                        byte=values.at(i).toLocal8Bit();
                        Text=byte.data();                        
                        check=sheet->writeStr(RowCount,11,Text);
                        break;
                    case IDENTITY:
                        Split=values.at(i).split("-");
                        byte=Split.at(0).toLocal8Bit();
                        Text=byte.data();
                        check=sheet->writeStr(RowCount,7,Text);
                        byte=Split.at(1).toLocal8Bit();
                        Text=byte.data();
                        check=sheet->writeStr(RowCount,8,Text);
                        break;
                    case VALVELINENUMBER:
                        if(values.at(i).toInt()>9)
                        {
                            Tmp=QString("%1").arg(values.at(i));
                        }
                        else
                        {
                            Tmp=QString("0%1").arg(values.at(i));
                        }
                        byte=Tmp.toLocal8Bit();
                        Text=byte.data();
                        check=sheet->writeStr(RowCount,10,Text);
                        break;
                    case VALVENUMBER:
                        Split=values.at(i).split("-");
                        byte=Split.at(0).toLocal8Bit();
                        Text=byte.data();
                        check=sheet->writeStr(RowCount,5,Text);
                        byte=Split.at(1).toLocal8Bit();
                        Text=byte.data();
                        check=sheet->writeStr(RowCount,6,Text);
                        break;
                    }
                }                
            }
            values.clear();
            RowCount++;
        }

        Split=FileName.split(".xls");
        Tmp=QString("%1_Convert.xls").arg(Split.at(0));        
        //qDebug()<<book->save(File_Name);
        byte=Tmp.toLocal8Bit();
        Text=byte.data();
        qDebug()<<book->save(Text);
        Statusbar_Show(Tmp+tr(" File Converted."));        
    }
    ValveBlock.clear();
    book->release();
}


void MainWindow::dragEnterEvent(QDragEnterEvent *e)
{
    if (e->mimeData()->hasUrls()) {
        e->acceptProposedAction();
    }
}

void MainWindow::dropEvent(QDropEvent *e)
{
    QStringList NameSplit;
    foreach (const QUrl &url, e->mimeData()->urls()) {
        QString fileName = url.toLocalFile();
        //url.toLocalFile()
        //qDebug() << "Dropped file:" << fileName;

        NameSplit=fileName.split(".");

        if(NameSplit.count()>1)
        {
            if(NameSplit.at(1)=="csv")
            {
                ui->listWidget_ValveFiles->addItem(fileName);
            }
            else if(NameSplit.at(1)=="xls")
            {
                ui->listWidget_DrawingFiles->addItem(fileName);
            }
        }
    }
}

void MainWindow::Statusbar_Show(QString Text)
{
    ui->statusBar->showMessage(Text);
}

void MainWindow::on_pushButton_Delete_clicked()
{
    QList<QListWidgetItem*> items = ui->listWidget_DrawingFiles->selectedItems();
    foreach(QListWidgetItem *item, items){
        delete ui->listWidget_DrawingFiles->takeItem(ui->listWidget_DrawingFiles->row(item));
    }

    items = ui->listWidget_ValveFiles->selectedItems();
    foreach(QListWidgetItem *item, items){
        delete ui->listWidget_ValveFiles->takeItem(ui->listWidget_ValveFiles->row(item));
    }    
}

void MainWindow::on_pushButton_Convert_clicked()
{
    ui->statusBar->clearMessage();
    if(ui->listWidget_DrawingFiles->count()==ui->listWidget_ValveFiles->count())
    {
        for(int i=0; i<ui->listWidget_ValveFiles->count(); i++)
        {
            ValveFileRead(ui->listWidget_ValveFiles->item(i)->text());
            DrawingFileWrite(ui->listWidget_DrawingFiles->item(i)->text());
        }

        if(ui->checkBox->isChecked())
        {
            QStringList Text=ui->listWidget_DrawingFiles->item(ui->listWidget_DrawingFiles->count()-1)->text().split("/");
            QString Tmp=ui->listWidget_DrawingFiles->item(ui->listWidget_DrawingFiles->count()-1)->text().replace(Text.at(Text.count()-1),"");
            QDesktopServices::openUrl(Tmp);
        }
    }

    else
    {
        Statusbar_Show(tr("ListCount is different!"));
    }
}

void MainWindow::on_pushButton_DeleteAll_clicked()
{
    ui->listWidget_DrawingFiles->clear();
    ui->listWidget_ValveFiles->clear();
}

void MainWindow::on_actionHelp_triggered()
{
    /*Process=new QProcess(this);
    qDebug()<<QApplication::applicationDirPath()+"/"+"help.xlsx";
    Process->start(QApplication::applicationDirPath()+"/"+"help.xlsx");

    if(Process!=NULL)
    {
       if(Process->isOpen())
       {
           Process->close();
       }
       Process->deleteLater();
    }*/

    QString newfilename = QApplication::applicationDirPath()+"/"+"help.pdf";
    QDesktopServices::openUrl(QUrl(newfilename.prepend( "file:///" )));
}

void MainWindow::on_pushButton_Open_clicked()
{
    QFileDialog dialog(this);
    dialog.setDirectory("C:/Users/user/AppData/Local/Temp");
    dialog.setFileMode(QFileDialog::ExistingFiles);
    dialog.setNameFilter(trUtf8("(*.csv *.xls)"));
    QStringList fileNames,Filter;
    if(dialog.exec())
    {
        fileNames = dialog.selectedFiles();

        foreach (QString FileName, fileNames)
        {
            Filter=FileName.split(".");

            if(Filter.count()>0)
            {
                if(Filter.at(1)=="csv")
                {
                    ui->listWidget_ValveFiles->addItem(FileName);
                }
                else if(Filter.at(1)=="xls")
                {
                    ui->listWidget_DrawingFiles->addItem(FileName);
                }
            }
        }
    }
}
