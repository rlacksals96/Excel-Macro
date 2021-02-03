package me.prj;

import javax.swing.*;
import javax.swing.filechooser.FileNameExtensionFilter;
import java.awt.*;
import java.awt.event.ActionEvent;
import java.awt.event.ActionListener;
import java.io.File;
import java.io.IOException;
import java.util.ArrayList;

public class MyView{
    private DefaultListModel<String> model;
    ArrayList<String> fileList;
    File []files;

    void setLayout(){
        JFrame frame=new JFrame("Macro Rules!!");
        frame.setDefaultCloseOperation(WindowConstants.DISPOSE_ON_CLOSE);
        JScrollPane scrollPane=new JScrollPane();
        JList<String> openList=new JList();
        JLabel title=new JLabel();
        JButton deleteBtn=new JButton("Delete");
        deleteBtn.setEnabled(false);


//        JLabel filePath=new JLabel("select path to save.");
        JTextArea fileNameInput=new JTextArea();
        JLabel filePathLabel=new JLabel("path:");
        JButton filePathBtn=new JButton("select");
        JLabel fileNameLabel=new JLabel("name:");
        JLabel xlsLabel=new JLabel(".xlsx");
        JButton buildBtn=new JButton("build");

//        JLabel notice=new JLabel("after click build, wait until notice. It takes about 5 sec per file");
        JLabel notice=new JLabel("<html>after click build, wait until notice!<br>It takes about 6 sec per file</html>");
//        buildBtn.setEnabled(false);
        JLabel fileCnt=new JLabel("uploaded files: 0");
        title.setText("file name format ex) 000000_00ㅇㅇㅇ(ㅁㅁㅁ)");
        title.setFont(new Font("arial",Font.PLAIN,14));
        title.setBounds(10,10,400,30);
        frame.add(title);

        JButton addBtn=new JButton("add");
        addBtn.setBounds(10,50,100,50);
        addBtn.setFont(new Font("arial",Font.PLAIN,17));
        addBtn.addActionListener(new ActionListener() {
            @Override
            public void actionPerformed(ActionEvent e) {
                JFileChooser chooser=new JFileChooser();
                chooser.setMultiSelectionEnabled(true);
                chooser.setFileFilter(new FileNameExtensionFilter("xlsx", "xlsx", "xls","xlsm"));
                int returnVal=chooser.showOpenDialog(null);
                if(returnVal == JFileChooser.APPROVE_OPTION) {
                    files = chooser.getSelectedFiles();
                    fileList=new ArrayList<String>();

                    try {
                        for(File file:files){
                            String fileName=file.getName();

//                            int idx = fileName .lastIndexOf(".");
//                            String fileName_=fileName.substring(0,idx);
                            fileList.add(fileName);
                        }

//                      for(String str:fileList)
//                          System.out.println(str);
                        model=new DefaultListModel<String>();
                        for(String str:fileList)
                            model.addElement(str);
                        if(model.getSize()!=0)
                            deleteBtn.setEnabled(true);
                        openList.setModel(model);

                        fileCnt.setText("uploaded files: "+model.getSize());
                    }catch(Exception ex) {
                        ex.printStackTrace();
                    }

                }
                else
                {
                    System.out.println("cancel file open.");
                }

            }
        });



        deleteBtn.setBounds(150,50,100,50);
        deleteBtn.setFont(new Font("arial",Font.PLAIN,17));

        deleteBtn.addActionListener(new ActionListener() {
            @Override
            public void actionPerformed(ActionEvent e) {
                int index=openList.getSelectedIndex();
                model.remove(index);
                if(model.getSize()==0)
                    deleteBtn.setEnabled(false);
                openList.setModel(model);
                fileCnt.setText("uploaded files: "+model.getSize());

            }
        });
//        filePathBtn.addActionListener(new ActionListener() {
//            @Override
//            public void actionPerformed(ActionEvent e) {
//                JFileChooser pathSelector=new JFileChooser();
//                pathSelector.setMultiSelectionEnabled(false);
//                pathSelector.setFileSelectionMode(JFileChooser.DIRECTORIES_ONLY);
//                int returnVal=pathSelector.showOpenDialog(null);
//                if(returnVal==JFileChooser.APPROVE_OPTION) {
//                    File file = pathSelector.getSelectedFile();
//                    filePath.setText(file.getAbsolutePath());
//                }
//
//            }
//        });
        buildBtn.addActionListener(new ActionListener() {
            @Override
            public void actionPerformed(ActionEvent e) {
                try{
                    if(model.getSize()==0){
                        JOptionPane.showMessageDialog(null,"no uploaded files");
                    }
                    else{
                        ExcelController excel=new ExcelController(files,fileList);
                        //excel.mkExcel();
                        long time1=System.currentTimeMillis();
                        try {
                            excel.editFile();
                        } catch (IOException ioException) {
                            ioException.printStackTrace();
                        }
                        long time2=System.currentTimeMillis();
                        double result=(time2-time1)/1000.0;
                        JOptionPane.showMessageDialog(null,"build success\nexec time: "+result);

                    }
                }catch(NullPointerException np){
                    JOptionPane.showMessageDialog(null,"no uploaded files");
                    np.printStackTrace();


                }

            }
        });

        filePathLabel.setBounds(280,120,50,30);
        filePathLabel.setFont(new Font("arial",Font.PLAIN,17));
//        filePathBtn.setBounds(330,120,70,30);
//        filePathBtn.setFont(new Font("arial",Font.PLAIN,13));
//        filePath.setBounds(410,120,250,30);
//        filePath.setFont(new Font("arial",Font.PLAIN,13));

        fileNameLabel.setBounds(280,160,50,30);
        fileNameLabel.setFont(new Font("arial",Font.PLAIN,17));
        fileNameInput.setBounds(340,165,180,20);
        xlsLabel.setBounds(530,165,40,20);

        buildBtn.setBounds(280,200,250,80);
        buildBtn.setFont(new Font("arial",Font.PLAIN,17));

        notice.setBounds(290,300,300,50);
        notice.setFont(new Font("arial",Font.BOLD,14));
        notice.setForeground(Color.BLUE);

        scrollPane.setBounds(10,120,250,330);
        scrollPane.setViewportView(openList);
//        openList.setBounds(10,120,250,330);
        fileCnt.setBounds(10,460,200,20);

        frame.add(notice);
        frame.add(fileCnt);
        frame.add(buildBtn);
//        frame.add(fileNameInput);
//        frame.add(xlsLabel);
//        frame.add(filePath);
//        frame.add(filePathBtn);
//        frame.add(filePathLabel);
//        frame.add(fileNameLabel);

        frame.add(scrollPane);
//        frame.add(openList);
        frame.add(addBtn);
        frame.add(deleteBtn);
        frame.setBounds(20,40,650,520);
        frame.setLayout(null);
        frame.setVisible(true);

    }



}
