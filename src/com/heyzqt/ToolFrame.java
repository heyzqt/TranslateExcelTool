package com.heyzqt;

import com.widget.FileChooser;
import com.widget.ToolFont;

import javax.swing.*;
import java.awt.*;
import java.awt.event.ActionEvent;
import java.awt.event.ActionListener;

/**
 * Created by heyzqt 9/25/2017
 */
public class ToolFrame extends JFrame implements ActionListener {

    /**
     * main frame
     */
    private JFrame mFrame;

    /**
     * main panel
     */
    private JPanel mMainPanel;

    /**
     * choose excel panel
     */
    private JPanel mChooseExcelPanel;
    private JButton mChooseExcelBtn;
    private JLabel mChooseExcelLab;

    /**
     * choose country panel
     */
    private JPanel mChooseCountryPanel;

    /**
     * Excel operations
     */
    private JPanel mExcelPanel;
    private JPanel mIndexPanel;
    private JButton mRemoveBtn;
    private JButton mInsertBtn;
    private JButton mCopyCellBtn;
    private JButton mCopyColAToColBBtn;

    private JPanel mExcel2XMLPanel;

    /**
     * log panel
     */
    private JScrollPane mLogScrollPane;

    public ToolFrame() {
        initFrame();
    }

    private void initFrame() {
        mFrame = new JFrame(Constant.FRAME_TITLE + "_" + Constant.TOOL_VERSION + "_" + Constant.TOOL_DEVELOPER);

        mMainPanel = new JPanel(new GridLayout(5, 1));

        //init choose excel panel
        mChooseExcelPanel = new JPanel(new BorderLayout(20, 10));
        JPanel panel_1_1 = new JPanel();
        JPanel panel_1_2 = new JPanel(new FlowLayout(FlowLayout.LEFT));
        mChooseExcelBtn = new JButton("选择Excel文件");
        mChooseExcelLab = new JLabel("文件路径：");
        mChooseExcelLab.setFont(new ToolFont());
        mChooseExcelBtn.setFont(new ToolFont());
        mChooseExcelBtn.addActionListener(this);
        panel_1_1.add(mChooseExcelBtn);
        panel_1_2.add(mChooseExcelLab);
        mChooseExcelPanel.add(panel_1_1, BorderLayout.WEST);
        mChooseExcelPanel.add(panel_1_2, BorderLayout.CENTER);

        //init choose country panel
        mChooseCountryPanel = new JPanel(new BorderLayout());
        JPanel panel_2_1 = new JPanel(new FlowLayout());
        JPanel panel_2_2 = new JPanel(new FlowLayout());
        JLabel countryLab = new JLabel("请选择国家：");
        countryLab.setFont(new ToolFont());
        panel_2_1.add(countryLab);
        panel_2_1.setBackground(Color.YELLOW);
        panel_2_2.setBackground(Color.CYAN);
        mChooseCountryPanel.add(panel_2_1, BorderLayout.NORTH);
        mChooseCountryPanel.add(panel_2_2, BorderLayout.CENTER);

        //init Excel operations
        mExcelPanel = new JPanel(new BorderLayout());
        JPanel panel_3_1 = new JPanel();
        GridLayout gridLayout = new GridLayout(1, 4);
        gridLayout.setHgap(50);
        gridLayout.setVgap(50);
        JPanel panel_3_2 = new JPanel(gridLayout);
        JLabel excelLab = new JLabel("请选择需要进行的操作：");
        countryLab.setFont(new ToolFont());
        mRemoveBtn = new JButton("删除行数");
        mRemoveBtn.setFont(new ToolFont());
        mInsertBtn = new JButton("添加行数");
        mInsertBtn.setFont(new ToolFont());
        mCopyCellBtn = new JButton("复制单元格");
        mCopyCellBtn.setFont(new ToolFont());
        mCopyColAToColBBtn = new JButton("复制列");
        mCopyColAToColBBtn.setFont(new ToolFont());
        //mIndexPanel = new JPanel(new GridLayout(2, 4));
        //mIndexPanel.add();
        mIndexPanel = new JPanel();
        mIndexPanel.setBackground(Color.red);

        panel_3_1.add(excelLab);
        panel_3_2.add(mRemoveBtn);
        panel_3_2.add(mInsertBtn);
        panel_3_2.add(mCopyCellBtn);
        panel_3_2.add(mCopyColAToColBBtn);
        panel_3_1.setBackground(Color.PINK);
        panel_3_2.setBackground(Color.orange);
        mExcelPanel.add(panel_3_1, BorderLayout.NORTH);
        mExcelPanel.add(panel_3_2, BorderLayout.CENTER);
        mExcelPanel.add(mIndexPanel, BorderLayout.SOUTH);

        //init transform Excel to XML file panel
        mExcel2XMLPanel = new JPanel();
        mExcel2XMLPanel.setBackground(Color.black);

        //init log panel
        mLogScrollPane = new JScrollPane();
        mLogScrollPane.setBackground(Color.LIGHT_GRAY);


        mMainPanel.add(mChooseExcelPanel);
        mMainPanel.add(mChooseCountryPanel);
        mMainPanel.add(mExcelPanel);
        mMainPanel.add(mExcel2XMLPanel);
        mMainPanel.add(mLogScrollPane);

        mFrame.setSize(960, 540);
        mFrame.setDefaultCloseOperation(JFrame.EXIT_ON_CLOSE);
        mFrame.setLocationRelativeTo(null);
        mFrame.setVisible(true);
        mFrame.add(mMainPanel);
    }

    @Override
    public void actionPerformed(ActionEvent e) {
        String button = e.getActionCommand();
        if (button.equals("选择Excel文件")) {
            FileChooser fileChooser = new FileChooser();
        }
    }
}
