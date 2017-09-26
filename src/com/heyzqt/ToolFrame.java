package com.heyzqt;

import com.sun.scenario.effect.impl.sw.sse.SSEBlend_SRC_OUTPeer;
import com.widget.DefaultFont;
import com.widget.FileChooser;
import com.widget.ToolFont;

import javax.swing.*;
import java.awt.*;
import java.awt.event.ActionEvent;
import java.awt.event.ActionListener;
import java.awt.event.ItemEvent;
import java.awt.event.ItemListener;
import java.util.Arrays;

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
    private JCheckBox checkBox1AR;
    private JCheckBox checkBox2BG_rBG;
    private JCheckBox checkBox3CA;
    private JCheckBox checkBox4CS;
    private JCheckBox checkBox5CY;
    private JCheckBox checkBox6DA;
    private JCheckBox checkBox7DE;
    private JCheckBox checkBox8EL_rGR;
    private JCheckBox checkBox9ES;
    private JCheckBox checkBox10ES_rPR;
    private JCheckBox checkBox11ET;
    private JCheckBox checkBox12EU;
    private JCheckBox checkBox13FA_rIR;
    private JCheckBox checkBox14FI;
    private JCheckBox checkBox15FR;
    private JCheckBox checkBox16GD;
    private JCheckBox checkBox17GL;
    private JCheckBox checkBox18HR;
    private JCheckBox checkBox19HU;
    private JCheckBox checkBox20IN_rID;
    private JCheckBox checkBox21IT;
    private JCheckBox checkBox22IW_rIL;
    private JCheckBox checkBox23KK_rKZ;
    private JCheckBox checkBox24MN_rMN;
    private JCheckBox checkBox25MS_rMY;
    private JCheckBox checkBox26MY_rMM;
    private JCheckBox checkBox27NB;
    private JCheckBox checkBox28NL;
    private JCheckBox checkBox29PL;
    private JCheckBox checkBox30PT;
    private JCheckBox checkBox31RO;
    private JCheckBox checkBox32RU;
    private JCheckBox checkBox33SK;
    private JCheckBox checkBox34SL;
    private JCheckBox checkBox35SQ_rAL;
    private JCheckBox checkBox36SR;
    private JCheckBox checkBox37SV;
    private JCheckBox checkBox38SW_rTZ;
    private JCheckBox checkBox39TA_rIN;
    private JCheckBox checkBox40TH;
    private JCheckBox checkBox41TR;
    private JCheckBox checkBox42UK_rUA;
    private JCheckBox checkBox43VI_rVN;
    private JCheckBox checkBox44ZH_rCN;
    private JCheckBox checkBox45ZH_rHK;
    private JCheckBox checkBox46ZH_rTW;


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
    private JTextField removeField1;
    private JTextField removeField2;
    private JTextField removeField3;
    private JTextField removeField4;
    private JTextField insertField1;
    private JTextField insertField2;
    private JTextField insertField3;
    private JTextField insertField4;
    private JTextField cpCellField1;
    private JTextField cpCellField2;
    private JTextField cpCellField3;
    private JTextField cpCellField4;
    private JTextField cpCellField5;
    private JTextField cpCellField6;
    private JTextField cpCellField7;
    private JTextField cpColField1;
    private JTextField cpColField2;
    private JTextField cpColField3;
    private JTextField cpColField4;
    private JTextField cpColField5;

    /**
     * transform Excel to XML file
     */
    private JPanel mExcel2XMLPanel;
    private JTextField transformField1;
    private JTextField transformField2;
    private JTextField transformField3;
    private JTextField transformField4;
    private JTextField transformField5;
    private JTextField transformField6;
    private String mFileType = Constant.FILE_STRINGS;
    public static int count = 0;

    /**
     * log panel
     */
    private JScrollPane mLogScrollPane;
    public static JTextArea mLogArea;

    private FileChooser mFileChooser;

    private String FILEPATH = "";

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
        JPanel panel4_1 = new JPanel(new BorderLayout());
        JButton clearLogBtn = new JButton("clear log");
        clearLogBtn.setFont(new ToolFont());
        clearLogBtn.addActionListener(this);
        panel4_1.add(mLogScrollPane, BorderLayout.CENTER);
        JPanel temp = new JPanel(new FlowLayout());
        temp.add(clearLogBtn);
        panel4_1.add(temp, BorderLayout.SOUTH);
        mMainPanel.add(panel4_1);

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
        checkBox1AR = new JCheckBox("ar");
        checkBox2BG_rBG = new JCheckBox("bg-rBG");
        checkBox3CA = new JCheckBox("ca");
        checkBox4CS = new JCheckBox("cs");
        checkBox5CY = new JCheckBox("cy");
        checkBox6DA = new JCheckBox("da");
        checkBox7DE = new JCheckBox("de");
        checkBox8EL_rGR = new JCheckBox("el-rGR");
        checkBox9ES = new JCheckBox("es");
        checkBox10ES_rPR = new JCheckBox("es-rPR");
        checkBox11ET = new JCheckBox("et");
        checkBox12EU = new JCheckBox("eu");
        checkBox13FA_rIR = new JCheckBox("fa-rIR");
        checkBox14FI = new JCheckBox("fi");
        checkBox15FR = new JCheckBox("fr");
        checkBox16GD = new JCheckBox("gd");
        checkBox17GL = new JCheckBox("gl");
        checkBox18HR = new JCheckBox("hr");
        checkBox19HU = new JCheckBox("hu");
        checkBox20IN_rID = new JCheckBox("in-rID");
        checkBox21IT = new JCheckBox("it");
        checkBox22IW_rIL = new JCheckBox("iw-rIL");
        checkBox23KK_rKZ = new JCheckBox("kk-rKZ");
        checkBox24MN_rMN = new JCheckBox("mn-rMN");
        checkBox25MS_rMY = new JCheckBox("ms-rMY");
        checkBox26MY_rMM = new JCheckBox("my-rMM");
        checkBox27NB = new JCheckBox("nb");
        checkBox28NL = new JCheckBox("nl");
        checkBox29PL = new JCheckBox("pl");
        checkBox30PT = new JCheckBox("pt");
        checkBox31RO = new JCheckBox("ro");
        checkBox32RU = new JCheckBox("ru");
        checkBox33SK = new JCheckBox("sk");
        checkBox34SL = new JCheckBox("sl");
        checkBox35SQ_rAL = new JCheckBox("sq-rAL");
        checkBox36SR = new JCheckBox("sr");
        checkBox37SV = new JCheckBox("sv");
        checkBox38SW_rTZ = new JCheckBox("sw-rTZ");
        checkBox39TA_rIN = new JCheckBox("ta-rIN");
        checkBox40TH = new JCheckBox("th");
        checkBox41TR = new JCheckBox("tr");
        checkBox42UK_rUA = new JCheckBox("uk-rUA");
        checkBox43VI_rVN = new JCheckBox("vi-rVN");
        checkBox44ZH_rCN = new JCheckBox("zh-rCN");
        checkBox45ZH_rHK = new JCheckBox("zh-rHK");
        checkBox46ZH_rTW = new JCheckBox("zh-rTW");

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
        checkBox24MN_rMN.setFont(new ToolFont());
        checkBox25MS_rMY.setFont(new ToolFont());
        checkBox26MY_rMM.setFont(new ToolFont());
        checkBox27NB.setFont(new ToolFont());
        checkBox28NL.setFont(new ToolFont());
        checkBox29PL.setFont(new ToolFont());
        checkBox30PT.setFont(new ToolFont());
        checkBox31RO.setFont(new ToolFont());
        checkBox32RU.setFont(new ToolFont());
        checkBox33SK.setFont(new ToolFont());
        checkBox34SL.setFont(new ToolFont());
        checkBox35SQ_rAL.setFont(new ToolFont());
        checkBox36SR.setFont(new ToolFont());
        checkBox37SV.setFont(new ToolFont());
        checkBox38SW_rTZ.setFont(new ToolFont());
        checkBox39TA_rIN.setFont(new ToolFont());
        checkBox40TH.setFont(new ToolFont());
        checkBox41TR.setFont(new ToolFont());
        checkBox42UK_rUA.setFont(new ToolFont());
        checkBox43VI_rVN.setFont(new ToolFont());
        checkBox44ZH_rCN.setFont(new ToolFont());
        checkBox45ZH_rHK.setFont(new ToolFont());
        checkBox46ZH_rTW.setFont(new ToolFont());

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
        checkBox24MN_rMN.addItemListener(this);
        checkBox25MS_rMY.addItemListener(this);
        checkBox26MY_rMM.addItemListener(this);
        checkBox27NB.addItemListener(this);
        checkBox28NL.addItemListener(this);
        checkBox29PL.addItemListener(this);
        checkBox30PT.addItemListener(this);
        checkBox31RO.addItemListener(this);
        checkBox32RU.addItemListener(this);
        checkBox33SK.addItemListener(this);
        checkBox34SL.addItemListener(this);
        checkBox35SQ_rAL.addItemListener(this);
        checkBox36SR.addItemListener(this);
        checkBox37SV.addItemListener(this);
        checkBox38SW_rTZ.addItemListener(this);
        checkBox39TA_rIN.addItemListener(this);
        checkBox40TH.addItemListener(this);
        checkBox41TR.addItemListener(this);
        checkBox42UK_rUA.addItemListener(this);
        checkBox43VI_rVN.addItemListener(this);
        checkBox44ZH_rCN.addItemListener(this);
        checkBox45ZH_rHK.addItemListener(this);
        checkBox46ZH_rTW.addItemListener(this);

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
        panel_2_2.add(checkBox24MN_rMN);
        panel_2_2.add(checkBox25MS_rMY);
        panel_2_2.add(checkBox26MY_rMM);
        panel_2_2.add(checkBox27NB);
        panel_2_2.add(checkBox28NL);
        panel_2_2.add(checkBox29PL);
        panel_2_2.add(checkBox30PT);
        panel_2_2.add(checkBox31RO);
        panel_2_2.add(checkBox32RU);
        panel_2_2.add(checkBox33SK);
        panel_2_2.add(checkBox34SL);
        panel_2_2.add(checkBox35SQ_rAL);
        panel_2_2.add(checkBox36SR);
        panel_2_2.add(checkBox37SV);
        panel_2_2.add(checkBox38SW_rTZ);
        panel_2_2.add(checkBox39TA_rIN);
        panel_2_2.add(checkBox40TH);
        panel_2_2.add(checkBox41TR);
        panel_2_2.add(checkBox42UK_rUA);
        panel_2_2.add(checkBox43VI_rVN);
        panel_2_2.add(checkBox44ZH_rCN);
        panel_2_2.add(checkBox45ZH_rHK);
        panel_2_2.add(checkBox46ZH_rTW);

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
        removeField1 = new JTextField(15);
        removeField2 = new JTextField(15);
        removeField3 = new JTextField(15);
        removeField4 = new JTextField(15);
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

        mRemoveConfirmBtn = new JButton("确认删除");
        mRemoveConfirmBtn.addActionListener(this);
        mRemoveConfirmBtn.setFont(new ToolFont());
        JPanel confirmPanel = new JPanel(new FlowLayout(FlowLayout.CENTER));
        confirmPanel.add(mRemoveConfirmBtn);
        removePal.add(confirmPanel);

        //insert panel
        JLabel insertLab1 = new JLabel("开始表序号（1、2、3...）：");
        JLabel insertLab2 = new JLabel("结束表序号（1、2、3...）：");
        JLabel insertLab3 = new JLabel("开始插入行序号(1、2、3...)：");
        JLabel insertLab4 = new JLabel("要插入多少行：");
        insertLab1.setFont(new ToolFont());
        insertLab2.setFont(new ToolFont());
        insertLab3.setFont(new ToolFont());
        insertLab4.setFont(new ToolFont());
        insertField1 = new JTextField(15);
        insertField2 = new JTextField(15);
        insertField3 = new JTextField(15);
        insertField4 = new JTextField(15);
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

        mInsertConfirmBtn = new JButton("确认插入");
        mInsertConfirmBtn.setFont(new ToolFont());
        mInsertConfirmBtn.addActionListener(this);
        JPanel insertPanel = new JPanel(new FlowLayout(FlowLayout.CENTER));
        insertPanel.add(mInsertConfirmBtn);
        insertPal.add(insertPanel);

        //copy cell panel
        JLabel cpCellLab1 = new JLabel("开始表序号（1、2、3...）：");
        JLabel cpCellLab2 = new JLabel("结束表序号（1、2、3...）：");
        JLabel cpCellLab3 = new JLabel("读取单元格列序号(1、2、3...)：");
        JLabel cpCellLab4 = new JLabel("读取单元格行序号(1、2、3...)：");
        JLabel cpCellLab5 = new JLabel("写入单元格列序号(1、2、3...)：");
        JLabel cpCellLab6 = new JLabel("写入单元格行序号(1、2、3...)：");
        JLabel cpCellLab7 = new JLabel("写入多少行（1、2、3...）：");
        cpCellLab1.setFont(new ToolFont());
        cpCellLab2.setFont(new ToolFont());
        cpCellLab3.setFont(new ToolFont());
        cpCellLab4.setFont(new ToolFont());
        cpCellLab5.setFont(new ToolFont());
        cpCellLab6.setFont(new ToolFont());
        cpCellLab7.setFont(new ToolFont());
        cpCellField1 = new JTextField(15);
        cpCellField2 = new JTextField(15);
        cpCellField3 = new JTextField(15);
        cpCellField4 = new JTextField(15);
        cpCellField5 = new JTextField(15);
        cpCellField6 = new JTextField(15);
        cpCellField7 = new JTextField(15);
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

        mCpCellConfirmBtn = new JButton("确认复制单元格");
        mCpCellConfirmBtn.setFont(new ToolFont());
        mCpCellConfirmBtn.addActionListener(this);
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
        cpColField1 = new JTextField(15);
        cpColField2 = new JTextField(15);
        cpColField3 = new JTextField(15);
        cpColField4 = new JTextField(15);
        cpColField5 = new JTextField(15);
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

        mCpColConfirmBtn = new JButton("确认复制列");
        mCpColConfirmBtn.setFont(new ToolFont());
        mCpColConfirmBtn.addActionListener(this);
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
        JRadioButton stringsBtn = new JRadioButton(Constant.FILE_STRINGS, true);
        JRadioButton menu_stringsBtn = new JRadioButton(Constant.FILE_MENU_STRINGS, false);
        JRadioButton nav_stringsBtn = new JRadioButton(Constant.FILE_NAV_STRINGS, false);
        JRadioButton cec_stringsBtn = new JRadioButton(Constant.FILE_CEC_STRINGS, false);
        JRadioButton mmp_stringsBtn = new JRadioButton(Constant.FILE_MMP_STRINGS, false);
        JRadioButton thr_menu_stringsBtn = new JRadioButton(Constant.FILE_THR_MENU_STRINGS, false);
        JRadioButton timeshift_stringsBtn = new JRadioButton(Constant.FILE_TIMESHIFT_STRINGS, false);

        stringsBtn.setFont(new ToolFont());
        menu_stringsBtn.setFont(new ToolFont());
        nav_stringsBtn.setFont(new ToolFont());
        cec_stringsBtn.setFont(new ToolFont());
        mmp_stringsBtn.setFont(new ToolFont());
        thr_menu_stringsBtn.setFont(new ToolFont());
        timeshift_stringsBtn.setFont(new ToolFont());

        stringsBtn.addItemListener(new RadioButtonListener());
        menu_stringsBtn.addItemListener(new RadioButtonListener());
        nav_stringsBtn.addItemListener(new RadioButtonListener());
        cec_stringsBtn.addItemListener(new RadioButtonListener());
        mmp_stringsBtn.addItemListener(new RadioButtonListener());
        thr_menu_stringsBtn.addItemListener(new RadioButtonListener());
        timeshift_stringsBtn.addItemListener(new RadioButtonListener());

        fileTypeGroup.add(stringsBtn);
        fileTypeGroup.add(menu_stringsBtn);
        fileTypeGroup.add(nav_stringsBtn);
        fileTypeGroup.add(cec_stringsBtn);
        fileTypeGroup.add(mmp_stringsBtn);
        fileTypeGroup.add(thr_menu_stringsBtn);
        fileTypeGroup.add(timeshift_stringsBtn);
        mExcel2XMLPanel.add(stringsBtn);
        mExcel2XMLPanel.add(menu_stringsBtn);
        mExcel2XMLPanel.add(nav_stringsBtn);
        mExcel2XMLPanel.add(cec_stringsBtn);
        mExcel2XMLPanel.add(mmp_stringsBtn);
        mExcel2XMLPanel.add(thr_menu_stringsBtn);
        mExcel2XMLPanel.add(timeshift_stringsBtn);

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
        transformField1 = new JTextField(15);
        transformField2 = new JTextField(15);
        transformField3 = new JTextField(15);
        transformField4 = new JTextField(15);
        transformField5 = new JTextField(15);
        transformField6 = new JTextField(15);
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
        transConfirmBtn.addActionListener(this);
        JPanel confirmPanel = new JPanel(new FlowLayout(FlowLayout.CENTER));
        confirmPanel.add(transConfirmBtn);
        mExcel2XMLPanel.add(confirmPanel);
    }

    private void initLogPanel() {
        //init log panel
        mLogArea = new JTextArea();
        mLogArea.setFont(new DefaultFont());
        mLogArea.append("this is a log area");
        mLogScrollPane = new JScrollPane(mLogArea);
    }

    @Override
    public void actionPerformed(ActionEvent e) {
        String btn = e.getActionCommand();
        if (btn.equals("选择Excel文件")) {
            System.out.println("choose Excel file");
            mFileChooser = new FileChooser();
            FILEPATH = mFileChooser.getFilepath();
            showLog(FILEPATH);
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
        } else if ("确认删除".equals(btn)) {
            int beginSheetIndex = 0;
            int endSheetIndex = 0;
            int beginRow = 0;
            int endRow = 0;

            try {
                beginSheetIndex = Integer.parseInt(removeField1.getText().trim());
                endSheetIndex = Integer.parseInt(removeField2.getText().trim());
                beginRow = Integer.parseInt(removeField3.getText().trim());
                endRow = Integer.parseInt(removeField4.getText().trim());
            } catch (NumberFormatException e1) {
                showLog("警告！！！参数填写有误，请检查后重新输入。");
                return;
            }

            if (FILEPATH.equals("")) {
                showLog("警告！！！还未选择文件。");
                return;
            }
            Main.removeRow(FILEPATH, beginSheetIndex, endSheetIndex, beginRow, endRow);
        } else if ("确认插入".equals(btn)) {
            int beginSheetIndex = 0;
            int endSheetIndex = 0;
            int beginRow = 0;
            int lines = 0;

            try {
                beginSheetIndex = Integer.parseInt(insertField1.getText().trim());
                endSheetIndex = Integer.parseInt(insertField2.getText().trim());
                beginRow = Integer.parseInt(insertField3.getText().trim());
                lines = Integer.parseInt(insertField4.getText().trim());
            } catch (NumberFormatException e1) {
                showLog("警告！！！参数填写有误，请检查后重新输入。");
                return;
            }

            if (FILEPATH.equals("")) {
                showLog("警告！！！还未选择文件。");
                return;
            }
            Main.insertRow(FILEPATH, beginSheetIndex, endSheetIndex, beginRow, lines);
        } else if ("确认复制单元格".equals(btn)) {
            int beginSheetIndex = 0;
            int endSheetIndex = 0;
            int readCol = 0;
            int readRow = 0;
            int writeCol = 0;
            int writeRow = 0;
            int lines = 0;

            try {
                beginSheetIndex = Integer.parseInt(cpCellField1.getText().trim());
                endSheetIndex = Integer.parseInt(cpCellField2.getText().trim());
                readCol = Integer.parseInt(cpCellField3.getText().trim());
                readRow = Integer.parseInt(cpCellField4.getText().trim());
                writeCol = Integer.parseInt(cpCellField5.getText().trim());
                writeRow = Integer.parseInt(cpCellField6.getText().trim());
                lines = Integer.parseInt(cpCellField7.getText().trim());
            } catch (NumberFormatException e1) {
                showLog("警告！！！参数填写有误，请检查后重新输入。");
                return;
            }

            if (FILEPATH.equals("")) {
                showLog("警告！！！还未选择文件。");
                return;
            }
            Main.addSingleCell(FILEPATH, beginSheetIndex, endSheetIndex,
                    readCol, readRow, writeCol, writeRow, lines);
        } else if ("确认复制列".equals(btn)) {
            int readSheetIndex = 0;
            int beginSheetIndex = 0;
            int endSheetIndex = 0;
            int readCol = 0;
            int writeCol = 0;

            try {
                readSheetIndex = Integer.parseInt(cpColField1.getText().trim());
                beginSheetIndex = Integer.parseInt(cpColField2.getText().trim());
                endSheetIndex = Integer.parseInt(cpColField3.getText().trim());
                readCol = Integer.parseInt(cpColField4.getText().trim());
                writeCol = Integer.parseInt(cpColField5.getText().trim());
            } catch (NumberFormatException e1) {
                showLog("警告！！！参数填写有误，请检查后重新输入。");
                return;
            }

            if (FILEPATH.equals("")) {
                showLog("警告！！！还未选择文件。");
                return;
            }
            Main.copyRowA2RowB(FILEPATH, readSheetIndex, beginSheetIndex, endSheetIndex, readCol, writeCol);
        } else if ("开始转换".equals(btn)) {

            if (count == 0) {
                showLog("警告！！！请选择国家");
                return;
            }

            //create file name
            showLog("The number of selected countries is " + count);
            showLog("selected file type is " + mFileType);
            String[] filenames = new String[count];
            String[] countries;
            countries = selectedCountries(count);
            showLog("choose countries = " + Arrays.toString(countries));
            switch (mFileType) {
                case Constant.FILE_STRINGS:
                    filenames = Main.createFileNames(Constant.FILE_STRINGS, countries, count);
                    break;
                case Constant.FILE_MENU_STRINGS:
                    filenames = Main.createFileNames(Constant.FILE_MENU_STRINGS, countries, count);
                    break;
                case Constant.FILE_NAV_STRINGS:
                    filenames = Main.createFileNames(Constant.FILE_NAV_STRINGS, countries, count);
                    break;
                case Constant.FILE_TIMESHIFT_STRINGS:
                    filenames = Main.createFileNames(Constant.FILE_TIMESHIFT_STRINGS, countries, count);
                    break;
                case Constant.FILE_MMP_STRINGS:
                    filenames = Main.createFileNames(Constant.FILE_MMP_STRINGS, countries, count);
                    break;
                case Constant.FILE_CEC_STRINGS:
                    filenames = Main.createFileNames(Constant.FILE_CEC_STRINGS, countries, count);
                    break;
                case Constant.FILE_THR_MENU_STRINGS:
                    filenames = Main.createFileNames(Constant.FILE_THR_MENU_STRINGS, countries, count);
                    break;
            }
            showLog("files name :  \n" + Arrays.toString(filenames));

//            int beginSheetIndex = 0;
//            int endSheetIndex = 0;
//            int keyCol = 0;
//            int valueCol = 0;
//            int beginRow = 0;
//            int endRow = 0;
//
//            try {
//                beginSheetIndex = Integer.parseInt(transformField1.getText().trim());
//                endSheetIndex = Integer.parseInt(transformField2.getText().trim());
//                keyCol = Integer.parseInt(transformField3.getText().trim());
//                valueCol = Integer.parseInt(transformField4.getText().trim());
//                beginRow = Integer.parseInt(transformField5.getText().trim());
//                endRow = Integer.parseInt(transformField6.getText().trim());
//            } catch (NumberFormatException e1) {
//                showLog("警告！！！参数填写有误，请检查后重新输入。");
//                return;
//            }
//
//            if (FILEPATH.equals("")) {
//                showLog("警告！！！还未选择文件。");
//                return;
//            }

//            Main.transformEXCEL2XML(FILEPATH, Main.XMLPATH, beginSheetIndex, endSheetIndex
//                    , keyCol, valueCol, beginRow, endRow);
        } else if ("clear log".equals(btn)) {
            mLogArea.setText("");
        }
    }

    private String[] selectedCountries(int count) {
        if (count == 0) {
            return null;
        }

        String[] result = new String[count];
        int index = 0;
        if (checkBox1AR.isSelected()) {
            result[index++] = checkBox1AR.getText().toString();
        }
        if (checkBox2BG_rBG.isSelected()) {
            result[index++] = checkBox2BG_rBG.getText().toString();
        }
        if (checkBox3CA.isSelected()) {
            result[index++] = checkBox3CA.getText().toString();
        }
        if (checkBox4CS.isSelected()) {
            result[index++] = checkBox4CS.getText().toString();
        }
        if (checkBox5CY.isSelected()) {
            result[index++] = checkBox5CY.getText().toString();
        }
        if (checkBox6DA.isSelected()) {
            result[index++] = checkBox6DA.getText().toString();
        }
        if (checkBox7DE.isSelected()) {
            result[index++] = checkBox7DE.getText().toString();
        }
        if (checkBox8EL_rGR.isSelected()) {
            result[index++] = checkBox8EL_rGR.getText().toString();
        }
        if (checkBox9ES.isSelected()) {
            result[index++] = checkBox9ES.getText().toString();
        }
        if (checkBox10ES_rPR.isSelected()) {
            result[index++] = checkBox10ES_rPR.getText().toString();
        }


        if (checkBox11ET.isSelected()) {
            result[index++] = checkBox11ET.getText().toString();
        }
        if (checkBox12EU.isSelected()) {
            result[index++] = checkBox12EU.getText().toString();
        }
        if (checkBox13FA_rIR.isSelected()) {
            result[index++] = checkBox13FA_rIR.getText().toString();
        }
        if (checkBox14FI.isSelected()) {
            result[index++] = checkBox14FI.getText().toString();
        }
        if (checkBox15FR.isSelected()) {
            result[index++] = checkBox15FR.getText().toString();
        }
        if (checkBox16GD.isSelected()) {
            result[index++] = checkBox16GD.getText().toString();
        }
        if (checkBox17GL.isSelected()) {
            result[index++] = checkBox17GL.getText().toString();
        }
        if (checkBox18HR.isSelected()) {
            result[index++] = checkBox18HR.getText().toString();
        }
        if (checkBox19HU.isSelected()) {
            result[index++] = checkBox19HU.getText().toString();
        }
        if (checkBox20IN_rID.isSelected()) {
            result[index++] = checkBox20IN_rID.getText().toString();
        }


        if (checkBox21IT.isSelected()) {
            result[index++] = checkBox21IT.getText().toString();
        }
        if (checkBox22IW_rIL.isSelected()) {
            result[index++] = checkBox22IW_rIL.getText().toString();
        }
        if (checkBox23KK_rKZ.isSelected()) {
            result[index++] = checkBox23KK_rKZ.getText().toString();
        }
        if (checkBox24MN_rMN.isSelected()) {
            result[index++] = checkBox24MN_rMN.getText().toString();
        }
        if (checkBox25MS_rMY.isSelected()) {
            result[index++] = checkBox25MS_rMY.getText().toString();
        }
        if (checkBox26MY_rMM.isSelected()) {
            result[index++] = checkBox26MY_rMM.getText().toString();
        }
        if (checkBox27NB.isSelected()) {
            result[index++] = checkBox27NB.getText().toString();
        }
        if (checkBox28NL.isSelected()) {
            result[index++] = checkBox28NL.getText().toString();
        }
        if (checkBox29PL.isSelected()) {
            result[index++] = checkBox29PL.getText().toString();
        }
        if (checkBox30PT.isSelected()) {
            result[index++] = checkBox30PT.getText().toString();
        }


        if (checkBox31RO.isSelected()) {
            result[index++] = checkBox31RO.getText().toString();
        }
        if (checkBox32RU.isSelected()) {
            result[index++] = checkBox32RU.getText().toString();
        }
        if (checkBox33SK.isSelected()) {
            result[index++] = checkBox33SK.getText().toString();
        }
        if (checkBox34SL.isSelected()) {
            result[index++] = checkBox34SL.getText().toString();
        }
        if (checkBox35SQ_rAL.isSelected()) {
            result[index++] = checkBox35SQ_rAL.getText().toString();
        }
        if (checkBox36SR.isSelected()) {
            result[index++] = checkBox36SR.getText().toString();
        }
        if (checkBox37SV.isSelected()) {
            result[index++] = checkBox37SV.getText().toString();
        }
        if (checkBox38SW_rTZ.isSelected()) {
            result[index++] = checkBox38SW_rTZ.getText().toString();
        }
        if (checkBox39TA_rIN.isSelected()) {
            result[index++] = checkBox39TA_rIN.getText().toString();
        }
        if (checkBox40TH.isSelected()) {
            result[index++] = checkBox40TH.getText().toString();
        }


        if (checkBox41TR.isSelected()) {
            result[index++] = checkBox41TR.getText().toString();
        }
        if (checkBox42UK_rUA.isSelected()) {
            result[index++] = checkBox42UK_rUA.getText().toString();
        }
        if (checkBox43VI_rVN.isSelected()) {
            result[index++] = checkBox43VI_rVN.getText().toString();
        }
        if (checkBox44ZH_rCN.isSelected()) {
            result[index++] = checkBox44ZH_rCN.getText().toString();
        }
        if (checkBox45ZH_rHK.isSelected()) {
            result[index++] = checkBox45ZH_rHK.getText().toString();
        }
        if (checkBox46ZH_rTW.isSelected()) {
            result[index++] = checkBox46ZH_rTW.getText().toString();
        }
        return result;
    }

    @Override
    public void itemStateChanged(ItemEvent e) {
        JCheckBox jcb = (JCheckBox) e.getItem();
        if (jcb.isSelected()) {
            System.out.println(jcb.getText() + " check");
            count++;
        } else {
            System.out.println(jcb.getText() + "check not");
            count--;
        }
    }

    class RadioButtonListener implements ItemListener {

        @Override
        public void itemStateChanged(ItemEvent e) {
            JRadioButton jrb = (JRadioButton) e.getSource();
            if (jrb.isSelected()) {
                mLogArea.append("\n" + jrb.getText().toString() + " is choosed.");
                mFileType = jrb.getText().toString();
            }
        }
    }

    public static void showLog(String msg) {
        mLogArea.append("\n" + msg);
    }
}
