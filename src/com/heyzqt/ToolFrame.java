package com.heyzqt;

import javax.swing.*;
import java.awt.*;

/**
 * Created by heyzqt 9/25/2017
 */
public class ToolFrame extends JFrame {

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
    private JLabel mChooseExcelLabel;

    private JPanel mChooseCountryPanel;
    private JPanel mExcelPanel;
    private JPanel mExcel2XMLPanel;
    private JPanel mLogPanel;


    public ToolFrame() {
        initFrame();
    }

    private void initFrame() {
        mFrame = new JFrame(Constant.FRAME_TITLE + "_" + Constant.TOOL_VERSION + "_" + Constant.TOOL_DEVELOPER);
        mFrame.setSize(960, 540);
        mFrame.setDefaultCloseOperation(JFrame.EXIT_ON_CLOSE);
        mFrame.setLocationRelativeTo(null);
        mFrame.setVisible(true);

        mMainPanel = new JPanel(new GridLayout(4, 1));

        //init choose excel panel
        mChooseExcelPanel = new JPanel(new FlowLayout());
        mChooseExcelBtn = new JButton("选择Excel文件");

        mChooseExcelPanel.add(new Label("hello"));

        mChooseCountryPanel = new JPanel(new FlowLayout());
        mChooseCountryPanel.add(new Label("world"));

        mMainPanel.add(mChooseExcelPanel);
        mMainPanel.add(mChooseCountryPanel);
        mFrame.add(mMainPanel);
    }
}
