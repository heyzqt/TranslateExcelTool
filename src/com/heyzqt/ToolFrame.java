package com.heyzqt;

import com.widget.FileChooser;
import com.widget.ToolFont;

import javax.swing.*;
import java.awt.*;
import java.awt.event.ActionEvent;
import java.awt.event.ActionListener;
import java.awt.event.ItemEvent;
import java.awt.event.ItemListener;

/**
 * Created by heyzqt 9/25/2017
 */
public class ToolFrame extends JFrame implements ActionListener, ItemListener {

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
    private JPanel mInfoPanel;
    private JPanel mIndexCardPanel;
    private JButton mRemoveBtn;
    private JButton mInsertBtn;
    private JButton mCopyCellBtn;
    private JButton mCopyColAToColBBtn;
    private JLabel mInfoLab;
    private JPanel removePal;
    private JPanel insertPal;
    private JPanel cpCellPal;
    private JPanel cpColPal;
    private CardLayout mOperationsCard;
    private JButton mRemoveConfirmBtn;
    private JButton mInsertConfirmBtn;
    private JButton mCpCellConfirmBtn;
    private JButton mCpColConfirmBtn;


    private JPanel mExcel2XMLPanel;

    /**
     * log panel
     */
    private JScrollPane mLogScrollPane;
    public static JTextArea mLogArea;

    private FileChooser mFileChooser;

    private String FILEPATH = "";
    private int mBeginSheetIndex;
    private int mEndSheetIndex;
    private int mBeginRow;
    private int mEndRow;

    public ToolFrame() {
        initFrame();
    }

    private void initFrame() {
        mFrame = new JFrame(Constant.FRAME_TITLE + "_" + Constant.TOOL_VERSION + "_" + Constant.TOOL_DEVELOPER);

        mMainPanel = new JPanel(new GridLayout(4, 1));

        initChooseExcelPanel();

        initChooseCountryPanel();

        initExcelOperationsPanel();

        initTransformPanel();

        initLogPanel();

        JPanel panel_1 = new JPanel(new BorderLayout());
        panel_1.add(mChooseExcelPanel, BorderLayout.WEST);
        panel_1.add(mChooseCountryPanel, BorderLayout.CENTER);
        mMainPanel.add(panel_1);
        mMainPanel.add(mExcelPanel);
        mMainPanel.add(mExcel2XMLPanel);
        mMainPanel.add(mLogScrollPane);

        mFrame.setSize(1280, 960);
        mFrame.setDefaultCloseOperation(JFrame.EXIT_ON_CLOSE);
        mFrame.setLocationRelativeTo(null);
        mFrame.setVisible(true);
        mFrame.add(mMainPanel);
    }

    private void initChooseExcelPanel() {
        //init choose Excel panel
        mChooseExcelPanel = new JPanel(new BorderLayout());
        JPanel panel_1_1 = new JPanel(new FlowLayout(FlowLayout.LEFT));
        JPanel panel_1_2 = new JPanel(new FlowLayout(FlowLayout.LEFT));
        mChooseExcelBtn = new JButton("选择Excel文件");
        mChooseExcelLab = new JLabel("文件路径：");
        mChooseExcelLab.setFont(new ToolFont());
        mChooseExcelBtn.setFont(new ToolFont());
        mChooseExcelBtn.addActionListener(this);
        panel_1_1.add(mChooseExcelBtn);
        panel_1_2.add(mChooseExcelLab);
        mChooseExcelPanel.add(panel_1_1, BorderLayout.NORTH);
        mChooseExcelPanel.add(panel_1_2, BorderLayout.CENTER);
    }

    private void initChooseCountryPanel() {
        //init choose country panel
        mChooseCountryPanel = new JPanel(new BorderLayout());
        JPanel panel_2_1 = new JPanel(new FlowLayout(FlowLayout.LEFT));
        JPanel panel_2_2 = new JPanel(new GridLayout(6, 10));
        JLabel countryLab = new JLabel("请选择国家：");
        countryLab.setFont(new ToolFont());
        panel_2_1.add(countryLab);
        JCheckBox checkBox1AR = new JCheckBox("ar");
        JCheckBox checkBox2BG_rBG = new JCheckBox("bg-rBG");
        JCheckBox checkBox3CA = new JCheckBox("ca");
        JCheckBox checkBox4CS = new JCheckBox("cs");
        JCheckBox checkBox5CY = new JCheckBox("cy");
        JCheckBox checkBox6DA = new JCheckBox("da");
        JCheckBox checkBox7DE = new JCheckBox("de");
        JCheckBox checkBox8EL_rGR = new JCheckBox("el-rGR");
        JCheckBox checkBox9ES = new JCheckBox("es");
        JCheckBox checkBox10ES_rPR = new JCheckBox("es-rPR");
        JCheckBox checkBox11ET = new JCheckBox("et");
        JCheckBox checkBox12EU = new JCheckBox("eu");
        JCheckBox checkBox13FA_rIR = new JCheckBox("fa-rIR");
        JCheckBox checkBox14FI = new JCheckBox("fi");
        JCheckBox checkBox15FR = new JCheckBox("fr");
        JCheckBox checkBox16GD = new JCheckBox("gd");
        JCheckBox checkBox17GL = new JCheckBox("gl");
        JCheckBox checkBox18HR = new JCheckBox("hr");
        JCheckBox checkBox19HU = new JCheckBox("hu");
        JCheckBox checkBox20IN_rID = new JCheckBox("in-rID");
        JCheckBox checkBox21IT = new JCheckBox("it");
        JCheckBox checkBox22IW_rIL = new JCheckBox("iw-rIL");
        JCheckBox checkBox23KK_rKZ = new JCheckBox("kk-rKZ");
        JCheckBox checkBox24LAND = new JCheckBox("land");
        JCheckBox checkBox25MN_rMN = new JCheckBox("mn-rMN");
        JCheckBox checkBox26MS_rMY = new JCheckBox("ms-rMY");
        JCheckBox checkBox27MY_rMM = new JCheckBox("my-rMM");
        JCheckBox checkBox28NB = new JCheckBox("nb");
        JCheckBox checkBox29NL = new JCheckBox("nl");
        JCheckBox checkBox30PL = new JCheckBox("pl");
        JCheckBox checkBox31PT = new JCheckBox("pt");
        JCheckBox checkBox32RO = new JCheckBox("ro");
        JCheckBox checkBox33RU = new JCheckBox("ru");
        JCheckBox checkBox34SK = new JCheckBox("sk");
        JCheckBox checkBox35SL = new JCheckBox("sl");
        JCheckBox checkBox36SQ_rAL = new JCheckBox("sq-rAL");
        JCheckBox checkBox37SR = new JCheckBox("sr");
        JCheckBox checkBox38SV = new JCheckBox("sv");
        JCheckBox checkBox39SW_rTZ = new JCheckBox("sw-rTZ");
        JCheckBox checkBox40TA_rIN = new JCheckBox("ta-rIN");
        JCheckBox checkBox41TH = new JCheckBox("th");
        JCheckBox checkBox42TR = new JCheckBox("tr");
        JCheckBox checkBox43UK_rUA = new JCheckBox("uk-rUA");
        JCheckBox checkBox44VI_rVN = new JCheckBox("vi-rVN");
        JCheckBox checkBox45ZH_rCN = new JCheckBox("zh-rCN");
        JCheckBox checkBox46ZH_rHK = new JCheckBox("zh-rHK");
        JCheckBox checkBox47ZH_rTW = new JCheckBox("zh-rTW");

        checkBox1AR.setFont(new ToolFont());
        checkBox2BG_rBG.setFont(new ToolFont());
        checkBox3CA.setFont(new ToolFont());
        checkBox4CS.setFont(new ToolFont());
        checkBox5CY.setFont(new ToolFont());
        checkBox6DA.setFont(new ToolFont());
        checkBox7DE.setFont(new ToolFont());
        checkBox8EL_rGR.setFont(new ToolFont());
        checkBox9ES.setFont(new ToolFont());
        checkBox10ES_rPR.setFont(new ToolFont());
        checkBox11ET.setFont(new ToolFont());
        checkBox12EU.setFont(new ToolFont());
        checkBox13FA_rIR.setFont(new ToolFont());
        checkBox14FI.setFont(new ToolFont());
        checkBox15FR.setFont(new ToolFont());
        checkBox16GD.setFont(new ToolFont());
        checkBox17GL.setFont(new ToolFont());
        checkBox18HR.setFont(new ToolFont());
        checkBox19HU.setFont(new ToolFont());
        checkBox20IN_rID.setFont(new ToolFont());
        checkBox21IT.setFont(new ToolFont());
        checkBox22IW_rIL.setFont(new ToolFont());
        checkBox23KK_rKZ.setFont(new ToolFont());
        checkBox24LAND.setFont(new ToolFont());
        checkBox25MN_rMN.setFont(new ToolFont());
        checkBox26MS_rMY.setFont(new ToolFont());
        checkBox27MY_rMM.setFont(new ToolFont());
        checkBox28NB.setFont(new ToolFont());
        checkBox29NL.setFont(new ToolFont());
        checkBox30PL.setFont(new ToolFont());
        checkBox31PT.setFont(new ToolFont());
        checkBox32RO.setFont(new ToolFont());
        checkBox33RU.setFont(new ToolFont());
        checkBox34SK.setFont(new ToolFont());
        checkBox35SL.setFont(new ToolFont());
        checkBox36SQ_rAL.setFont(new ToolFont());
        checkBox37SR.setFont(new ToolFont());
        checkBox38SV.setFont(new ToolFont());
        checkBox39SW_rTZ.setFont(new ToolFont());
        checkBox40TA_rIN.setFont(new ToolFont());
        checkBox41TH.setFont(new ToolFont());
        checkBox42TR.setFont(new ToolFont());
        checkBox43UK_rUA.setFont(new ToolFont());
        checkBox44VI_rVN.setFont(new ToolFont());
        checkBox45ZH_rCN.setFont(new ToolFont());
        checkBox46ZH_rHK.setFont(new ToolFont());
        checkBox47ZH_rTW.setFont(new ToolFont());

        checkBox1AR.addItemListener(this);
        checkBox2BG_rBG.addItemListener(this);
        checkBox3CA.addItemListener(this);
        checkBox4CS.addItemListener(this);
        checkBox5CY.addItemListener(this);
        checkBox6DA.addItemListener(this);
        checkBox7DE.addItemListener(this);
        checkBox8EL_rGR.addItemListener(this);
        checkBox9ES.addItemListener(this);
        checkBox10ES_rPR.addItemListener(this);
        checkBox11ET.addItemListener(this);
        checkBox12EU.addItemListener(this);
        checkBox13FA_rIR.addItemListener(this);
        checkBox14FI.addItemListener(this);
        checkBox15FR.addItemListener(this);
        checkBox16GD.addItemListener(this);
        checkBox17GL.addItemListener(this);
        checkBox18HR.addItemListener(this);
        checkBox19HU.addItemListener(this);
        checkBox20IN_rID.addItemListener(this);
        checkBox21IT.addItemListener(this);
        checkBox22IW_rIL.addItemListener(this);
        checkBox23KK_rKZ.addItemListener(this);
        checkBox24LAND.addItemListener(this);
        checkBox25MN_rMN.addItemListener(this);
        checkBox26MS_rMY.addItemListener(this);
        checkBox27MY_rMM.addItemListener(this);
        checkBox28NB.addItemListener(this);
        checkBox29NL.addItemListener(this);
        checkBox30PL.addItemListener(this);
        checkBox31PT.addItemListener(this);
        checkBox32RO.addItemListener(this);
        checkBox33RU.addItemListener(this);
        checkBox34SK.addItemListener(this);
        checkBox35SL.addItemListener(this);
        checkBox36SQ_rAL.addItemListener(this);
        checkBox37SR.addItemListener(this);
        checkBox38SV.addItemListener(this);
        checkBox39SW_rTZ.addItemListener(this);
        checkBox40TA_rIN.addItemListener(this);
        checkBox41TH.addItemListener(this);
        checkBox42TR.addItemListener(this);
        checkBox43UK_rUA.addItemListener(this);
        checkBox44VI_rVN.addItemListener(this);
        checkBox45ZH_rCN.addItemListener(this);
        checkBox46ZH_rHK.addItemListener(this);
        checkBox47ZH_rTW.addItemListener(this);

        panel_2_2.add(checkBox1AR);
        panel_2_2.add(checkBox2BG_rBG);
        panel_2_2.add(checkBox3CA);
        panel_2_2.add(checkBox4CS);
        panel_2_2.add(checkBox5CY);
        panel_2_2.add(checkBox6DA);
        panel_2_2.add(checkBox7DE);
        panel_2_2.add(checkBox8EL_rGR);
        panel_2_2.add(checkBox9ES);
        panel_2_2.add(checkBox10ES_rPR);
        panel_2_2.add(checkBox11ET);
        panel_2_2.add(checkBox12EU);
        panel_2_2.add(checkBox13FA_rIR);
        panel_2_2.add(checkBox14FI);
        panel_2_2.add(checkBox15FR);
        panel_2_2.add(checkBox16GD);
        panel_2_2.add(checkBox17GL);
        panel_2_2.add(checkBox18HR);
        panel_2_2.add(checkBox19HU);
        panel_2_2.add(checkBox20IN_rID);
        panel_2_2.add(checkBox21IT);
        panel_2_2.add(checkBox22IW_rIL);
        panel_2_2.add(checkBox23KK_rKZ);
        panel_2_2.add(checkBox24LAND);
        panel_2_2.add(checkBox25MN_rMN);
        panel_2_2.add(checkBox26MS_rMY);
        panel_2_2.add(checkBox27MY_rMM);
        panel_2_2.add(checkBox28NB);
        panel_2_2.add(checkBox29NL);
        panel_2_2.add(checkBox30PL);
        panel_2_2.add(checkBox31PT);
        panel_2_2.add(checkBox32RO);
        panel_2_2.add(checkBox33RU);
        panel_2_2.add(checkBox34SK);
        panel_2_2.add(checkBox35SL);
        panel_2_2.add(checkBox36SQ_rAL);
        panel_2_2.add(checkBox37SR);
        panel_2_2.add(checkBox38SV);
        panel_2_2.add(checkBox39SW_rTZ);
        panel_2_2.add(checkBox40TA_rIN);
        panel_2_2.add(checkBox41TH);
        panel_2_2.add(checkBox42TR);
        panel_2_2.add(checkBox43UK_rUA);
        panel_2_2.add(checkBox44VI_rVN);
        panel_2_2.add(checkBox45ZH_rCN);
        panel_2_2.add(checkBox46ZH_rHK);
        panel_2_2.add(checkBox47ZH_rTW);

        mChooseCountryPanel.add(panel_2_1, BorderLayout.NORTH);
        mChooseCountryPanel.add(panel_2_2, BorderLayout.CENTER);
    }

    private void initExcelOperationsPanel() {
        //init Excel operations
        mExcelPanel = new JPanel(new BorderLayout());
        JPanel panel_3_1 = new JPanel(new FlowLayout(FlowLayout.LEFT));
        JPanel panel_3_2 = new JPanel();
        JLabel excelLab = new JLabel("请选择需要进行的操作：");
        excelLab.setFont(new ToolFont());
        mRemoveBtn = new JButton("删除行数");
        mRemoveBtn.setFont(new ToolFont());
        mInsertBtn = new JButton("插入行数");
        mInsertBtn.setFont(new ToolFont());
        mCopyCellBtn = new JButton("复制单元格");
        mCopyCellBtn.setFont(new ToolFont());
        mCopyColAToColBBtn = new JButton("复制列");
        mCopyColAToColBBtn.setFont(new ToolFont());
        mRemoveBtn.addActionListener(this);
        mInsertBtn.addActionListener(this);
        mCopyCellBtn.addActionListener(this);
        mCopyColAToColBBtn.addActionListener(this);

        // input the paramaters
        mOperationsCard = new CardLayout();
        mIndexCardPanel = new JPanel(mOperationsCard);
        removePal = new JPanel(new GridLayout(1, 5));
        insertPal = new JPanel(new GridLayout(1, 5));
        cpCellPal = new JPanel(new GridLayout(2, 5));
        cpColPal = new JPanel(new GridLayout(2, 5));

        // remove panel
        JLabel removeLab1 = new JLabel("开始表序号（1、2、3...）：");
        JLabel removeLab2 = new JLabel("结束表序号（1、2、3...）：");
        JLabel removeLab3 = new JLabel("开始删除行序号(1、2、3...)：");
        JLabel removeLab4 = new JLabel("结束删除行序号(1、2、3...)：");
        removeLab1.setFont(new ToolFont());
        removeLab2.setFont(new ToolFont());
        removeLab3.setFont(new ToolFont());
        removeLab4.setFont(new ToolFont());
        JTextField removeField1 = new JTextField(15);
        JTextField removeField2 = new JTextField(15);
        JTextField removeField3 = new JTextField(15);
        JTextField removeField4 = new JTextField(15);
        JPanel removePal_1 = new JPanel();
        JPanel removePal_2 = new JPanel();
        JPanel removePal_3 = new JPanel();
        JPanel removePal_4 = new JPanel();
        removePal_1.add(removeLab1);
        removePal_1.add(removeField1);
        removePal_2.add(removeLab2);
        removePal_2.add(removeField2);
        removePal_3.add(removeLab3);
        removePal_3.add(removeField3);
        removePal_4.add(removeLab4);
        removePal_4.add(removeField4);
        removePal.add(removePal_1);
        removePal.add(removePal_2);
        removePal.add(removePal_3);
        removePal.add(removePal_4);

        mRemoveConfirmBtn = new JButton("删除确认执行");
        mRemoveConfirmBtn.setFont(new ToolFont());
        JPanel confirmPanel = new JPanel(new FlowLayout(FlowLayout.CENTER));
        confirmPanel.add(mRemoveConfirmBtn);
        removePal.add(confirmPanel);

        //insert panel
        JLabel insertLab1 = new JLabel("开始表序号（1、2、3...）：");
        JLabel insertLab2 = new JLabel("结束表序号（1、2、3...）：");
        JLabel insertLab3 = new JLabel("开始插入行序号(1、2、3...)：");
        JLabel insertLab4 = new JLabel("要插入的行数：");
        insertLab1.setFont(new ToolFont());
        insertLab2.setFont(new ToolFont());
        insertLab3.setFont(new ToolFont());
        insertLab4.setFont(new ToolFont());
        JTextField insertField1 = new JTextField(15);
        JTextField insertField2 = new JTextField(15);
        JTextField insertField3 = new JTextField(15);
        JTextField insertField4 = new JTextField(15);
        JPanel insertPal_1 = new JPanel();
        JPanel insertPal_2 = new JPanel();
        JPanel insertPal_3 = new JPanel();
        JPanel insertPal_4 = new JPanel();
        insertPal_1.add(insertLab1);
        insertPal_1.add(insertField1);
        insertPal_2.add(insertLab2);
        insertPal_2.add(insertField2);
        insertPal_3.add(insertLab3);
        insertPal_3.add(insertField3);
        insertPal_4.add(insertLab4);
        insertPal_4.add(insertField4);
        insertPal.add(insertPal_1);
        insertPal.add(insertPal_2);
        insertPal.add(insertPal_3);
        insertPal.add(insertPal_4);

        mInsertConfirmBtn = new JButton("添加确认执行");
        mInsertConfirmBtn.setFont(new ToolFont());
        JPanel insertPanel = new JPanel(new FlowLayout(FlowLayout.CENTER));
        insertPanel.add(mInsertConfirmBtn);
        insertPal.add(insertPanel);

        //copy cell panel
        JLabel cpCellLab1 = new JLabel("开始表序号（1、2、3...）：");
        JLabel cpCellLab2 = new JLabel("结束表序号（1、2、3...）：");
        JLabel cpCellLab3 = new JLabel("读取单元格列数(1、2、3...)：");
        JLabel cpCellLab4 = new JLabel("读取单元格行数(1、2、3...)：");
        JLabel cpCellLab5 = new JLabel("写入单元格列数(1、2、3...)：");
        JLabel cpCellLab6 = new JLabel("写入单元格行数(1、2、3...)：");
        JLabel cpCellLab7 = new JLabel("写入多少行（1、2、3...）：");
        cpCellLab1.setFont(new ToolFont());
        cpCellLab2.setFont(new ToolFont());
        cpCellLab3.setFont(new ToolFont());
        cpCellLab4.setFont(new ToolFont());
        cpCellLab5.setFont(new ToolFont());
        cpCellLab6.setFont(new ToolFont());
        cpCellLab7.setFont(new ToolFont());
        JTextField cpCellField1 = new JTextField(15);
        JTextField cpCellField2 = new JTextField(15);
        JTextField cpCellField3 = new JTextField(15);
        JTextField cpCellField4 = new JTextField(15);
        JTextField cpCellField5 = new JTextField(15);
        JTextField cpCellField6 = new JTextField(15);
        JTextField cpCellField7 = new JTextField(15);
        JPanel cpCellPal_1 = new JPanel();
        JPanel cpCellPal_2 = new JPanel();
        JPanel cpCellPal_3 = new JPanel();
        JPanel cpCellPal_4 = new JPanel();
        JPanel cpCellPal_5 = new JPanel();
        JPanel cpCellPal_6 = new JPanel();
        JPanel cpCellPal_7 = new JPanel();
        cpCellPal_1.add(cpCellLab1);
        cpCellPal_1.add(cpCellField1);
        cpCellPal_2.add(cpCellLab2);
        cpCellPal_2.add(cpCellField2);
        cpCellPal_3.add(cpCellLab3);
        cpCellPal_3.add(cpCellField3);
        cpCellPal_4.add(cpCellLab4);
        cpCellPal_4.add(cpCellField4);
        cpCellPal_5.add(cpCellLab5);
        cpCellPal_5.add(cpCellField5);
        cpCellPal_6.add(cpCellLab6);
        cpCellPal_6.add(cpCellField6);
        cpCellPal_7.add(cpCellLab7);
        cpCellPal_7.add(cpCellField7);
        cpCellPal.add(cpCellPal_1);
        cpCellPal.add(cpCellPal_2);
        cpCellPal.add(cpCellPal_3);
        cpCellPal.add(cpCellPal_4);
        cpCellPal.add(cpCellPal_5);
        cpCellPal.add(cpCellPal_6);
        cpCellPal.add(cpCellPal_7);

        mCpCellConfirmBtn = new JButton("复制单元格确认执行");
        mCpCellConfirmBtn.setFont(new ToolFont());
        JPanel cpCellPanel = new JPanel(new FlowLayout(FlowLayout.CENTER));
        cpCellPanel.add(mCpCellConfirmBtn);
        cpCellPal.add(cpCellPanel);

        //copy col panel
        JLabel cpColLab1 = new JLabel("读取表序号（1、2、3...）：");
        JLabel cpColLab2 = new JLabel("开始写入表序号（1、2、3...）：");
        JLabel cpColLab3 = new JLabel("结束写入表序号(1、2、3...)：");
        JLabel cpColLab4 = new JLabel("读取列序号(1、2、3...)：");
        JLabel cpColLab5 = new JLabel("写入列序号(1、2、3...)：");
        cpColLab1.setFont(new ToolFont());
        cpColLab2.setFont(new ToolFont());
        cpColLab3.setFont(new ToolFont());
        cpColLab4.setFont(new ToolFont());
        cpColLab5.setFont(new ToolFont());
        JTextField cpColField1 = new JTextField(15);
        JTextField cpColField2 = new JTextField(15);
        JTextField cpColField3 = new JTextField(15);
        JTextField cpColField4 = new JTextField(15);
        JTextField cpColField5 = new JTextField(15);
        JPanel cpColPal_1 = new JPanel();
        JPanel cpColPal_2 = new JPanel();
        JPanel cpColPal_3 = new JPanel();
        JPanel cpColPal_4 = new JPanel();
        JPanel cpColPal_5 = new JPanel();
        cpColPal_1.add(cpColLab1);
        cpColPal_1.add(cpColField1);
        cpColPal_2.add(cpColLab2);
        cpColPal_2.add(cpColField2);
        cpColPal_3.add(cpColLab3);
        cpColPal_3.add(cpColField3);
        cpColPal_4.add(cpColLab4);
        cpColPal_4.add(cpColField4);
        cpColPal_5.add(cpColLab5);
        cpColPal_5.add(cpColField5);
        cpColPal.add(cpColPal_1);
        cpColPal.add(cpColPal_2);
        cpColPal.add(cpColPal_3);
        cpColPal.add(cpColPal_4);
        cpColPal.add(cpColPal_5);

        mCpColConfirmBtn = new JButton("复制列确认执行");
        mCpColConfirmBtn.setFont(new ToolFont());
        JPanel cpColPanel = new JPanel(new FlowLayout(FlowLayout.CENTER));
        cpColPanel.add(mCpColConfirmBtn);
        cpColPal.add(cpColPanel);

        panel_3_1.add(excelLab);
        panel_3_2.add(mRemoveBtn);
        panel_3_2.add(mInsertBtn);
        panel_3_2.add(mCopyCellBtn);
        panel_3_2.add(mCopyColAToColBBtn);

        mIndexCardPanel.add("defaultcard", new JLabel());
        mIndexCardPanel.add("card1", removePal);
        mIndexCardPanel.add("card2", insertPal);
        mIndexCardPanel.add("card3", cpCellPal);
        mIndexCardPanel.add("card4", cpColPal);

        mInfoPanel = new JPanel(new BorderLayout());
        mInfoLab = new JLabel();
        mInfoLab.setFont(new ToolFont());
        mInfoPanel.add(mInfoLab);

        JPanel panel = new JPanel();
        panel.add(panel_3_1);
        panel.add(panel_3_2);
        mExcelPanel.add(panel, BorderLayout.NORTH);
        mExcelPanel.add(mIndexCardPanel, BorderLayout.CENTER);
        mExcelPanel.add(mInfoPanel, BorderLayout.SOUTH);
    }

    private void initTransformPanel() {
        //init transform Excel to XML file panel
        mExcel2XMLPanel = new JPanel(new GridLayout(3, 5));
        JLabel label = new JLabel("将Excel文件转换为XML文件（不支持array类型）：");
        label.setFont(new ToolFont());
        mExcel2XMLPanel.add(label);

        ButtonGroup fileTypeGroup = new ButtonGroup();
        JRadioButton stringsBtn = new JRadioButton("strings.xml", true);
        JRadioButton menu_stringsBtn = new JRadioButton("menu_strings.xml", false);
        JRadioButton nav_stringsBtn = new JRadioButton("nav_strings.xml", false);
        JRadioButton cec_stringsBtn = new JRadioButton("cec_strings.xml", false);
        JRadioButton mmp_stringsBtn = new JRadioButton("mmp_strings.xml", false);
        JRadioButton thr_menu_stringsBtn = new JRadioButton("thr_menu_strings.xml", false);
        JRadioButton timeshift_tringsBtn = new JRadioButton("timeshift_strings.xml", false);

        stringsBtn.setFont(new ToolFont());
        menu_stringsBtn.setFont(new ToolFont());
        nav_stringsBtn.setFont(new ToolFont());
        cec_stringsBtn.setFont(new ToolFont());
        mmp_stringsBtn.setFont(new ToolFont());
        thr_menu_stringsBtn.setFont(new ToolFont());
        timeshift_tringsBtn.setFont(new ToolFont());

        stringsBtn.addItemListener(new RadioButtonListener());
        menu_stringsBtn.addItemListener(new RadioButtonListener());
        nav_stringsBtn.addItemListener(new RadioButtonListener());
        cec_stringsBtn.addItemListener(new RadioButtonListener());
        mmp_stringsBtn.addItemListener(new RadioButtonListener());
        thr_menu_stringsBtn.addItemListener(new RadioButtonListener());
        timeshift_tringsBtn.addItemListener(new RadioButtonListener());

        fileTypeGroup.add(stringsBtn);
        fileTypeGroup.add(menu_stringsBtn);
        fileTypeGroup.add(nav_stringsBtn);
        fileTypeGroup.add(cec_stringsBtn);
        fileTypeGroup.add(mmp_stringsBtn);
        fileTypeGroup.add(thr_menu_stringsBtn);
        fileTypeGroup.add(timeshift_tringsBtn);
        mExcel2XMLPanel.add(stringsBtn);
        mExcel2XMLPanel.add(menu_stringsBtn);
        mExcel2XMLPanel.add(nav_stringsBtn);
        mExcel2XMLPanel.add(cec_stringsBtn);
        mExcel2XMLPanel.add(mmp_stringsBtn);
        mExcel2XMLPanel.add(thr_menu_stringsBtn);
        mExcel2XMLPanel.add(timeshift_tringsBtn);

        JLabel transformLab1 = new JLabel("开始表序号（1、2、3...）：");
        JLabel transformLab2 = new JLabel("结束表序号（1、2、3...）：");
        JLabel transformLab3 = new JLabel("key值列数序号(1、2、3...)：");
        JLabel transformLab4 = new JLabel("value值列数序号(1、2、3...)：");
        JLabel transformLab5 = new JLabel("开始写入行序号(1、2、3...)：");
        JLabel transformLab6 = new JLabel("结束写入行序号(1、2、3...)：");
        transformLab1.setFont(new ToolFont());
        transformLab2.setFont(new ToolFont());
        transformLab3.setFont(new ToolFont());
        transformLab4.setFont(new ToolFont());
        transformLab5.setFont(new ToolFont());
        transformLab6.setFont(new ToolFont());
        JTextField transformField1 = new JTextField(15);
        JTextField transformField2 = new JTextField(15);
        JTextField transformField3 = new JTextField(15);
        JTextField transformField4 = new JTextField(15);
        JTextField transformField5 = new JTextField(15);
        JTextField transformField6 = new JTextField(15);
        JPanel transformPal_1 = new JPanel();
        JPanel transformPal_2 = new JPanel();
        JPanel transformPal_3 = new JPanel();
        JPanel transformPal_4 = new JPanel();
        JPanel transformPal_5 = new JPanel();
        JPanel transformPal_6 = new JPanel();
        transformPal_1.add(transformLab1);
        transformPal_1.add(transformField1);
        transformPal_2.add(transformLab2);
        transformPal_2.add(transformField2);
        transformPal_3.add(transformLab3);
        transformPal_3.add(transformField3);
        transformPal_4.add(transformLab4);
        transformPal_4.add(transformField4);
        transformPal_5.add(transformLab5);
        transformPal_5.add(transformField5);
        transformPal_6.add(transformLab6);
        transformPal_6.add(transformField6);
        mExcel2XMLPanel.add(transformPal_1);
        mExcel2XMLPanel.add(transformPal_2);
        mExcel2XMLPanel.add(transformPal_3);
        mExcel2XMLPanel.add(transformPal_4);
        mExcel2XMLPanel.add(transformPal_5);
        mExcel2XMLPanel.add(transformPal_6);

        JButton transConfirmBtn = new JButton("开始转换");
        transConfirmBtn.setFont(new ToolFont());
        JPanel confirmPanel = new JPanel(new FlowLayout(FlowLayout.CENTER));
        confirmPanel.add(transConfirmBtn);
        mExcel2XMLPanel.add(confirmPanel);
    }

    private void initLogPanel() {
        //init log panel
        mLogArea = new JTextArea();
        mLogArea.setFont(new ToolFont());
        mLogArea.append("this is a log area");
        mLogScrollPane = new JScrollPane(mLogArea);
    }

    @Override
    public void actionPerformed(ActionEvent e) {
        String btn = e.getActionCommand();
        if (btn.equals("选择Excel文件")) {
            System.out.println("choose Excel");
            mFileChooser = new FileChooser();
        } else if (btn.equals("删除行数")) {
            mOperationsCard.show(mIndexCardPanel, "card1");
            mInfoLab.setText("删除行数说明：删除连续整行。比如：填入参数 1  10  2  5。意思是将表1到表10的第2行到第5行都删除");
        } else if (btn.equals("插入行数")) {
            mOperationsCard.show(mIndexCardPanel, "card2");
            mInfoLab.setText("插入行数说明：插入连续整行。比如：填入参数 1  10  2  3。意思是将表1到表10的从2行开始插入3行");
        } else if (btn.equals("复制单元格")) {
            mOperationsCard.show(mIndexCardPanel, "card3");
            mInfoLab.setText("复制单元格说明：读取每个表中指定单元格，然后写入每个表中指定的多个位置。\n");
        } else if (btn.equals("复制列")) {
            mOperationsCard.show(mIndexCardPanel, "card4");
            mInfoLab.setText("复制列说明：将表readSheetIndex的readColume列数据" +
                    "，复制到beginSheetIndex表到endSheetIndex的writeColume" +
                    "列中。\n比如：填入参数 1  2  10 3 3。意思是读取表1第3列的数据然后将第3列数据复制到表2到表10的第3列中");
        }
    }

    @Override
    public void itemStateChanged(ItemEvent e) {
        JCheckBox jcb = (JCheckBox) e.getItem();
        if (jcb.isSelected()) {
            System.out.println(jcb.getText() + " check");
        } else {
            System.out.println(jcb.getText() + "check not");
        }
    }

    class RadioButtonListener implements ItemListener {

        @Override
        public void itemStateChanged(ItemEvent e) {
            JRadioButton jrb = (JRadioButton) e.getSource();
            if (jrb.isSelected()) {
                mLogArea.append("\n" + jrb.getText() + " is choosed.");
            }
        }
    }
}
