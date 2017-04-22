/**
 * Created by vineet on 15-Apr-17.
 */

import javafx.application.Application;
import javafx.event.ActionEvent;
import javafx.event.EventHandler;
import javafx.geometry.HPos;
import javafx.geometry.Orientation;
import javafx.geometry.Pos;
import javafx.geometry.VPos;
import javafx.scene.Scene;
import javafx.scene.control.*;
import javafx.scene.image.Image;
import javafx.scene.layout.GridPane;
import javafx.scene.text.Text;
import javafx.stage.FileChooser;
import javafx.stage.Stage;
import net.objecthunter.exp4j.ExpressionBuilder;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFPrintSetup;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.poi.xwpf.usermodel.ParagraphAlignment;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFRun;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;

public class ExcelOutput extends Application {
    private static final String defaultFileName = "";
    private Stage savedStage;
    private ArrayList<Double> a1 = new ArrayList<>();

    private ArrayList<Double> ultimate = new ArrayList<>();
    private double sum, ultimateSum, ultimateGrandSum = 0;

    private int index = 1;
    private int indexX = 1;
    private int rowIndex = 5;
    private char subIndex = 'a';
    private char subIndexX = 'a';

    private TextField tankSize = new TextField();
    private TextField graniteSize = new TextField();

    private CheckBox ewCB = new CheckBox("E/W in excavation of soil");
    private CheckBox bfCB = new CheckBox("Supplying & providing 1st class B.F.");
    private CheckBox rcCB = new CheckBox("Supplying & providing R.C.C(1:2:4) with stone");
    private CheckBox thckBrkCB = new CheckBox("Supplying & providing 10” brick work in C.M.(1:6)");
    private CheckBox thnBrkCB = new CheckBox("Supplying & providing 5” B/W in C.M(1:5)");
    private CheckBox halfCB = new CheckBox("Supplying & providing 1\\2”C.P.in C.M(1:6)");
    private CheckBox qtrCB = new CheckBox("Supplying & providing 1\\4” C.P in C.M(1:4)");
    private CheckBox steelCB = new CheckBox("Supplying & providing different diameter of steel");
    private CheckBox tankCB = new CheckBox("Provision of septic tank");
    private CheckBox woodCB = new CheckBox("Supplying & providing with fixing well seasoned sal wood");
    private CheckBox ipsCB = new CheckBox("Supplying & Providing 1” I.P.S(1:2:4) with punning");
    private CheckBox graniteCB = new CheckBox("Granite On 3” P.C.C.(1:3:6)");
    private CheckBox paintCB = new CheckBox("Supplying & Providing Two coats of enamel paints to chaukhats & pannels");
    private CheckBox snowCB = new CheckBox("Supplying & Providing Two coats of snow cement to walls & ceiling ");
    private CheckBox xtraCB = new CheckBox("EXTRA SERVICES");

    //for options in R.C.C.
    private CheckBox aCB = new CheckBox("Pocket(main & support)");
    private CheckBox bCB = new CheckBox("10”x12” G.B.");
    private CheckBox cCB = new CheckBox("10”x10” Stiffeners");
    private CheckBox dCB = new CheckBox("10”x6” Lintel band");
    private CheckBox eCB = new CheckBox("Chajja, Loft etc.");
    private CheckBox fCB = new CheckBox("Roof slab");
    private CheckBox gCB = new CheckBox("Stair case");
    private CheckBox hCB = new CheckBox("Front Beam");
    private CheckBox iCB = new CheckBox("10”x10” beam on 5” B/W under roof slab");
    private CheckBox jCB = new CheckBox("Veranda beam");
    private CheckBox kCB = new CheckBox("R.C.C. railing through roof");

    //for options in extra part
    private CheckBox sanitaryCB = new CheckBox("Sanitary services");
    private CheckBox elecCB = new CheckBox("Electrical services");
    private CheckBox extraCB = new CheckBox("Extra services");
    private TextField totalPerc = new TextField();
    private Button ultimateGrandVal = new Button("CALCULATE GRAND TOTAL");
    private Text displayUltimateGrandVal = new Text();

    private boolean ewFlag, bfFlag, rcFlag, thckBrkFlag, thnBrkFlag, halfFlag, qtrFlag, steelFlag, tankFlag, woodFlag, ipsFlag, xtraFlag, graniteFlag, paintFlag, snowFlag = false;
    private boolean aFlag, bFlag, cFlag, dFlag, eFlag, fFlag, gFlag, hFlag, iFlag, jFlag, kFlag = false;
    private boolean sanitaryFlag, elecFlag, extraFlag = false;

    private TextField custName = new TextField();
    private TextArea custDetails = new TextArea();
    private TextField builtUp = new TextField();
    private TextField Feet = new TextField();
    private TextField inches = new TextField();
    private Button btConvert = new Button("Convert");
    private TextField totalFeet = new TextField();

    private TextField expression = new TextField();
    private Text result = new Text();
    private TextField rate = new TextField();
    private Button btCalculateQty = new Button("Calculate Qty.");
    private Button btAmount = new Button("Calculate Amt.");
    private Text tAmt = new Text();

    private TextField expression1 = new TextField();
    private Text result1 = new Text();
    private TextField rate1 = new TextField();
    private Button btCalculateQty1 = new Button("Calculate Qty.");
    private Button btAmount1 = new Button("Calculate Amt.");
    private Text tAmt1 = new Text();


    private Text prevRes = new Text();
    private Text result2 = new Text();
    private TextField sandRate2 = new TextField();
    private TextField rate2 = new TextField();
    private Button btCalculateQty2 = new Button("Calculate Qty.");
    private Button btAmount2 = new Button("Calculate Amt.");
    private Text tAmt2 = new Text();

    // the total sum of expressions
    private Button btSumTotal = new Button("Calculate Total sum");
    private Text totalSum = new Text();
    private Button btAmountSum = new Button("Calculate Amount");
    private TextField sumRate = new TextField();
    private Text tAmtSum = new Text();

    /**
     * this is for the R.C.C. optional part
     */
    private TextField expressionA = new TextField();
    private Button evalA = new Button("Evaluate");
    private Text resA = new Text();

    private TextField expressionB = new TextField();
    private Button evalB = new Button("Evaluate");
    private Text resB = new Text();

    private TextField expressionC = new TextField();
    private Button evalC = new Button("Evaluate");
    private Text resC = new Text();

    private TextField expressionD = new TextField();
    private Button evalD = new Button("Evaluate");
    private Text resD = new Text();

    private TextField expressionE = new TextField();
    private Button evalE = new Button("Evaluate");
    private Text resE = new Text();

    private TextField expressionF = new TextField();
    private Button evalF = new Button("Evaluate");
    private Text resF = new Text();

    private TextField expressionG = new TextField();
    private Button evalG = new Button("Evaluate");
    private Text resG = new Text();

    private TextField expressionH = new TextField();
    private Button evalH = new Button("Evaluate");
    private Text resH = new Text();

    private TextField expressionI = new TextField();
    private Button evalI = new Button("Evaluate");
    private Text resI = new Text();

    private TextField expressionJ = new TextField();
    private Button evalJ = new Button("Evaluate");
    private Text resJ = new Text();

    private TextField expressionK = new TextField();
    private Button evalK = new Button("Evaluate");
    private Text resK = new Text();

    /**
     * here ends the options for evaluation
     */

    private TextField expression3 = new TextField();
    private Text result3 = new Text();
    private TextField rate3 = new TextField();
    private Button btCalculateQty3 = new Button("Calculate Qty.");
    private Button btAmount3 = new Button("Calculate Amt.");
    private Text tAmt3 = new Text();

    private TextField expression4 = new TextField();
    private Text result4 = new Text();
    private TextField rate4 = new TextField();
    private Button btCalculateQty4 = new Button("Calculate Qty.");
    private Button btAmount4 = new Button("Calculate Amt.");
    private Text tAmt4 = new Text();

    private TextField expression5 = new TextField();
    private Text result5 = new Text();
    private TextField rate5 = new TextField();
    private Button btCalculateQty5 = new Button("Calculate Qty.");
    private Button btAmount5 = new Button("Calculate Amt.");
    private Text tAmt5 = new Text();

    private TextField expression6 = new TextField();
    private Text result6 = new Text();
    private TextField rate6 = new TextField();
    private Button btCalculateQty6 = new Button("Calculate Qty.");
    private Button btAmount6 = new Button("Calculate Amt.");
    private Text tAmt6 = new Text();

    private TextField expression7 = new TextField();
    private Text result7 = new Text();
    private TextField rate7 = new TextField();
    private Button btCalculateQty7 = new Button("Calculate Qty.");
    private Button btAmount7 = new Button("Calculate Amt.");
    private Text tAmt7 = new Text();

    //for septic tank
    private TextField expression8 = new TextField();
    private Button btAmount8 = new Button("Add to Amount");
    private Text tAmt8 = new Text();

    private TextField expression9 = new TextField();
    private Text result9 = new Text();
    private TextField rate9 = new TextField();
    private Button btCalculateQty9 = new Button("Calculate Qty.");
    private Button btAmount9 = new Button("Calculate Amt.");
    private Text tAmt9 = new Text();

    private TextField expression10 = new TextField();
    private Text result10 = new Text();
    private TextField rate10 = new TextField();
    private Button btCalculateQty10 = new Button("Calculate ");
    private Button btAmount10 = new Button("Calculate Amt.");
    private Text tAmt10 = new Text();

    private Text prevResA = new Text();
    private Text result2A = new Text();
    private TextField sandRate2A = new TextField();
    private TextField rate2A = new TextField();
    private Button btCalculateQty2A = new Button("Calculate Qty.");
    private Button btAmount2A = new Button("Calculate Amt.");
    private Text tAmt2A = new Text();

    private TextField expression11 = new TextField();
    private Text result11 = new Text();
    private TextField rate11 = new TextField();
    private Button btCalculateQty11 = new Button("Calculate Qty.");
    private Button btAmount11 = new Button("Calculate Amt.");
    private Text tAmt11 = new Text();

    private TextField expression12 = new TextField();
    private Text result12 = new Text();
    private TextField rate12 = new TextField();
    private Button btCalculateQty12 = new Button("Calculate Qty.");
    private Button btAmount12 = new Button("Calculate Amt.");
    private Text tAmt12 = new Text();

    private TextField expression13 = new TextField();
    private Text result13 = new Text();
    private TextField rate13 = new TextField();
    private Button btCalculateQty13 = new Button("Calculate Qty.");
    private Button btAmount13 = new Button("Calculate Amt.");
    private Text tAmt13 = new Text();

    private TextField elecPerc = new TextField();
    private TextField sanitaryPerc = new TextField();
    private TextField extraPerc = new TextField();

    private Button ultimateCal = new Button("Calculate cost of estimate without extra charges");
    private Text ultimateTarget = new Text();

    private Button generateX = new Button("Generate Excel file");
    private Button generate = new Button("Generate Word file");
    private Text actionTarget = new Text("");

    @Override // Override the start method in the Application class
    public void start(Stage primaryStage) {
        ScrollPane scrollPane = new ScrollPane();
        GridPane gridPane = new GridPane();

        // Create UI
        gridPane.setHgap(5);
        gridPane.setVgap(5);
        gridPane.add(new Label("Enter Customer Name"), 0, 0);
        gridPane.add(custName, 1, 0);

        gridPane.add(new Label("Enter other details"), 0, 1);
        gridPane.add(custDetails, 1, 1);

        gridPane.add(new Label("Enter built-up area in sft."), 0, 2);
        gridPane.add(builtUp, 1, 2);

        //vertical separator
        final Separator sepVert = new Separator();
        sepVert.setOrientation(Orientation.VERTICAL);
        sepVert.setValignment(VPos.CENTER);
        GridPane.setConstraints(sepVert, 2, 0);
        GridPane.setRowSpan(sepVert, 5);
        gridPane.getChildren().add(sepVert);

        gridPane.add(new Label("Feet To inches Calculator"), 3, 0);
        gridPane.add(new Label("Enter feet"), 3, 1);
        gridPane.add(Feet, 4, 1);
        gridPane.add(new Label("Enter inches"), 3, 2);
        gridPane.add(inches, 4, 2);
        gridPane.add(btConvert, 4, 3);
        gridPane.add(new Label("Total feet:"), 3, 4);
        gridPane.add(totalFeet, 4, 4);

        //1st horizontal separator
        final Separator sepHor1 = new Separator();
        sepHor1.setValignment(VPos.CENTER);
        GridPane.setConstraints(sepHor1, 0, 5);
        GridPane.setColumnSpan(sepHor1, 15);
        gridPane.getChildren().add(sepHor1);

        gridPane.add(ewCB, 0, 6);
        gridPane.add(new Label("Enter Expression:"), 0, 7);
        gridPane.add(expression, 1, 7);
        gridPane.add(new Label("Quantity(cft):"), 2, 7);
        gridPane.add(result, 3, 7);
        gridPane.add(new Label("Enter Rate (it is per 1000):"), 4, 7);
        gridPane.add(rate, 5, 7);
        gridPane.add(new Label("Amount:"), 6, 7);
        gridPane.add(tAmt, 7, 7);
        gridPane.add(btCalculateQty, 1, 8);
        gridPane.add(btAmount, 5, 8);

        //2nd horizontal separator
        final Separator sepHor2 = new Separator();
        sepHor2.setValignment(VPos.CENTER);
        GridPane.setConstraints(sepHor2, 0, 9);
        GridPane.setColumnSpan(sepHor2, 15);
        gridPane.getChildren().add(sepHor2);

        gridPane.add(bfCB, 0, 10);
        gridPane.add(new Label("Enter Expression:"), 0, 11);
        gridPane.add(expression1, 1, 11);
        gridPane.add(new Label("Quantity(sft):"), 2, 11);
        gridPane.add(result1, 3, 11);
        gridPane.add(new Label("Enter Rate % (it is per 100):"), 4, 11);
        gridPane.add(rate1, 5, 11);
        gridPane.add(new Label("Amount:"), 6, 11);
        gridPane.add(tAmt1, 7, 11);
        gridPane.add(btCalculateQty1, 1, 12);
        gridPane.add(btAmount1, 5, 12);

        //for local sand calculation
        gridPane.add(new Label("Local sand calculation"), 0, 13);
        gridPane.add(prevRes, 1, 13);
        gridPane.add(new Label("Enter inch(after converting to decimal)"), 2, 13);
        gridPane.add(sandRate2, 3, 13);
        gridPane.add(result2, 4, 13);
        gridPane.add(new Label("Enter Rate % (it is per 100):"), 5, 13);
        gridPane.add(rate2, 6, 13);
        gridPane.add(new Label("Amount:"), 7, 13);
        gridPane.add(tAmt2, 8, 13);
        gridPane.add(btCalculateQty2, 3, 14);
        gridPane.add(btAmount2, 6, 14);

        //3rd horizontal separator
        final Separator sepHor3 = new Separator();
        sepHor3.setValignment(VPos.CENTER);
        GridPane.setConstraints(sepHor3, 0, 15);
        GridPane.setColumnSpan(sepHor3, 15);
        gridPane.getChildren().add(sepHor3);

        gridPane.add(rcCB, 0, 16);

        /*for adding the options in rcc*/
        gridPane.add(aCB, 1, 16);
        gridPane.add(expressionA, 2, 16);
        gridPane.add(evalA, 3, 16);
        gridPane.add(new Label("Result is(cft):"), 4, 16);
        gridPane.add(resA, 5, 16);

        gridPane.add(bCB, 1, 17);
        gridPane.add(expressionB, 2, 17);
        gridPane.add(evalB, 3, 17);
        gridPane.add(new Label("Result is(cft):"), 4, 17);
        gridPane.add(resB, 5, 17);

        gridPane.add(cCB, 1, 18);
        gridPane.add(expressionC, 2, 18);
        gridPane.add(evalC, 3, 18);
        gridPane.add(new Label("Result is(cft):"), 4, 18);
        gridPane.add(resC, 5, 18);

        gridPane.add(dCB, 1, 19);
        gridPane.add(expressionD, 2, 19);
        gridPane.add(evalD, 3, 19);
        gridPane.add(new Label("Result is(cft):"), 4, 19);
        gridPane.add(resD, 5, 19);

        gridPane.add(eCB, 1, 20);
        gridPane.add(expressionE, 2, 20);
        gridPane.add(evalE, 3, 20);
        gridPane.add(new Label("Result is(cft):"), 4, 20);
        gridPane.add(resE, 5, 20);

        gridPane.add(fCB, 1, 21);
        gridPane.add(expressionF, 2, 21);
        gridPane.add(evalF, 3, 21);
        gridPane.add(new Label("Result is(cft):"), 4, 21);
        gridPane.add(resF, 5, 21);

        gridPane.add(gCB, 1, 22);
        gridPane.add(expressionG, 2, 22);
        gridPane.add(evalG, 3, 22);
        gridPane.add(new Label("Result is(cft):"), 4, 22);
        gridPane.add(resG, 5, 22);

        gridPane.add(hCB, 1, 23);
        gridPane.add(expressionH, 2, 23);
        gridPane.add(evalH, 3, 23);
        gridPane.add(new Label("Result is(cft):"), 4, 23);
        gridPane.add(resH, 5, 23);

        gridPane.add(iCB, 1, 24);
        gridPane.add(expressionI, 2, 24);
        gridPane.add(evalI, 3, 24);
        gridPane.add(new Label("Result is(cft):"), 4, 24);
        gridPane.add(resI, 5, 24);

        gridPane.add(jCB, 1, 25);
        gridPane.add(expressionJ, 2, 25);
        gridPane.add(evalJ, 3, 25);
        gridPane.add(new Label("Result is(cft):"), 4, 25);
        gridPane.add(resJ, 5, 25);

        gridPane.add(kCB, 1, 26);
        gridPane.add(expressionK, 2, 26);
        gridPane.add(evalK, 3, 26);
        gridPane.add(new Label("Result is(cft):"), 4, 26);
        gridPane.add(resK, 5, 26);

        gridPane.add(btSumTotal, 1, 27);
        gridPane.add(new Label("Total sum is(cft):"), 2, 27);
        gridPane.add(totalSum, 3, 27);
        gridPane.add(new Label("Enter Rate % (it is per 100):"), 4, 27);
        gridPane.add(sumRate, 5, 27);
        gridPane.add(new Label("Amount:"), 6, 27);
        gridPane.add(tAmtSum, 7, 27);
        gridPane.add(btAmountSum, 5, 28);

        //4th horizontal separator
        final Separator sepHor4 = new Separator();
        sepHor4.setValignment(VPos.CENTER);
        GridPane.setConstraints(sepHor4, 0, 29);
        GridPane.setColumnSpan(sepHor4, 15);
        gridPane.getChildren().add(sepHor4);

        gridPane.add(thckBrkCB, 0, 30);
        gridPane.add(new Label("Enter Expression:"), 0, 31);
        gridPane.add(expression3, 1, 31);
        gridPane.add(new Label("Quantity(cft):"), 2, 31);
        gridPane.add(result3, 3, 31);
        gridPane.add(new Label("Enter Rate % (it is per 100):"), 4, 31);
        gridPane.add(rate3, 5, 31);
        gridPane.add(new Label("Amount:"), 6, 31);
        gridPane.add(tAmt3, 7, 31);
        gridPane.add(btCalculateQty3, 1, 32);
        gridPane.add(btAmount3, 5, 32);

        //5th horizontal separator
        final Separator sepHor5 = new Separator();
        sepHor5.setValignment(VPos.CENTER);
        GridPane.setConstraints(sepHor5, 0, 33);
        GridPane.setColumnSpan(sepHor5, 15);
        gridPane.getChildren().add(sepHor5);

        gridPane.add(thnBrkCB, 0, 34);
        gridPane.add(new Label("Enter Expression:"), 0, 35);
        gridPane.add(expression4, 1, 35);
        gridPane.add(new Label("Quantity(cft):"), 2, 35);
        gridPane.add(result4, 3, 35);
        gridPane.add(new Label("Enter Rate % (it is per 100):"), 4, 35);
        gridPane.add(rate4, 5, 35);
        gridPane.add(new Label("Amount:"), 6, 35);
        gridPane.add(tAmt4, 7, 35);
        gridPane.add(btCalculateQty4, 1, 36);
        gridPane.add(btAmount4, 5, 36);

        //6th horizontal separator
        final Separator sepHor6 = new Separator();
        sepHor6.setValignment(VPos.CENTER);
        GridPane.setConstraints(sepHor6, 0, 37);
        GridPane.setColumnSpan(sepHor6, 15);
        gridPane.getChildren().add(sepHor6);

        gridPane.add(halfCB, 0, 38);
        gridPane.add(new Label("Enter Expression:"), 0, 39);
        gridPane.add(expression5, 1, 39);
        gridPane.add(new Label("Quantity(sft):"), 2, 39);
        gridPane.add(result5, 3, 39);
        gridPane.add(new Label("Enter Rate % (it is per 100):"), 4, 39);
        gridPane.add(rate5, 5, 39);
        gridPane.add(new Label("Amount:"), 6, 39);
        gridPane.add(tAmt5, 7, 39);
        gridPane.add(btCalculateQty5, 1, 40);
        gridPane.add(btAmount5, 5, 40);

        //7th horizontal separator
        final Separator sepHor7 = new Separator();
        sepHor7.setValignment(VPos.CENTER);
        GridPane.setConstraints(sepHor7, 0, 41);
        GridPane.setColumnSpan(sepHor7, 15);
        gridPane.getChildren().add(sepHor7);

        gridPane.add(qtrCB, 0, 42);
        gridPane.add(new Label("Enter Expression:"), 0, 43);
        gridPane.add(expression6, 1, 43);
        gridPane.add(new Label("Quantity(sft):"), 2, 43);
        gridPane.add(result6, 3, 43);
        gridPane.add(new Label("Enter Rate % (it is per 100):"), 4, 43);
        gridPane.add(rate6, 5, 43);
        gridPane.add(new Label("Amount:"), 6, 43);
        gridPane.add(tAmt6, 7, 43);
        gridPane.add(btCalculateQty6, 1, 44);
        gridPane.add(btAmount6, 5, 44);

        //8th horizontal separator
        final Separator sepHor8 = new Separator();
        sepHor8.setValignment(VPos.CENTER);
        GridPane.setConstraints(sepHor8, 0, 45);
        GridPane.setColumnSpan(sepHor8, 15);
        gridPane.getChildren().add(sepHor8);

        gridPane.add(steelCB, 0, 46);
        gridPane.add(new Label("Enter Expression:"), 0, 47);
        gridPane.add(expression7, 1, 47);
        gridPane.add(new Label("Quantity(in M.T.):"), 2, 47);
        gridPane.add(result7, 3, 47);
        gridPane.add(new Label("Enter Rate % (in P.M.T.):"), 4, 47);
        gridPane.add(rate7, 5, 47);
        gridPane.add(new Label("Amount:"), 6, 47);
        gridPane.add(tAmt7, 7, 47);
        gridPane.add(btCalculateQty7, 1, 48);
        gridPane.add(btAmount7, 5, 48);

        //9th horizontal separator
        final Separator sepHor9 = new Separator();
        sepHor9.setValignment(VPos.CENTER);
        GridPane.setConstraints(sepHor9, 0, 49);
        GridPane.setColumnSpan(sepHor9, 15);
        gridPane.getChildren().add(sepHor9);

        gridPane.add(tankCB, 0, 50);
        gridPane.add(new Label("Please enter dimensions of septic tank"), 1, 50);
        gridPane.add(tankSize, 2, 50);
        gridPane.add(new Label("Enter Lump Sum cost of tank"), 0, 51);
        gridPane.add(expression8, 1, 51);
        gridPane.add(new Label("Amount:"), 6, 51);
        gridPane.add(tAmt8, 7, 51);
        gridPane.add(btAmount8, 1, 52);

        //10th horizontal separator
        final Separator sepHor10 = new Separator();
        sepHor10.setValignment(VPos.CENTER);
        GridPane.setConstraints(sepHor10, 0, 53);
        GridPane.setColumnSpan(sepHor10, 15);
        gridPane.getChildren().add(sepHor10);

        gridPane.add(woodCB, 0, 54);
        gridPane.add(new Label("Enter Expression:"), 0, 55);
        gridPane.add(expression9, 1, 55);
        gridPane.add(new Label("Quantity(cft):"), 2, 55);
        gridPane.add(result9, 3, 55);
        gridPane.add(new Label("Enter Rate "), 4, 55);
        gridPane.add(rate9, 5, 55);
        gridPane.add(new Label("Amount:"), 6, 55);
        gridPane.add(tAmt9, 7, 55);
        gridPane.add(btCalculateQty9, 1, 56);
        gridPane.add(btAmount9, 5, 56);

        //11th horizontal separator
        final Separator sepHor11 = new Separator();
        sepHor11.setValignment(VPos.CENTER);
        GridPane.setConstraints(sepHor11, 0, 57);
        GridPane.setColumnSpan(sepHor11, 15);
        gridPane.getChildren().add(sepHor11);

        gridPane.add(ipsCB, 0, 58);
        gridPane.add(new Label("Enter Expression:"), 0, 59);
        gridPane.add(expression10, 1, 59);
        gridPane.add(new Label("Quantity(sft):"), 2, 59);
        gridPane.add(result10, 3, 59);
        gridPane.add(new Label("Enter Rate % (it is per 100):"), 4, 59);
        gridPane.add(rate10, 5, 59);
        gridPane.add(new Label("Amount:"), 6, 59);
        gridPane.add(tAmt10, 7, 59);
        gridPane.add(btCalculateQty10, 1, 60);
        gridPane.add(btAmount10, 5, 60);

        //for 2nd part of P.C.C. calculation
        gridPane.add(new Label(" 3” P.C.C.(1:3:6):"), 0, 61);
        gridPane.add(prevResA, 1, 61);
        gridPane.add(new Label("Enter inch(after converting to decimal)"), 2, 61);
        gridPane.add(sandRate2A, 3, 61);
        gridPane.add(result2A, 4, 61);
        gridPane.add(new Label("Enter Rate % (it is per 100):"), 5, 61);
        gridPane.add(rate2A, 6, 61);
        gridPane.add(new Label("Amount:"), 7, 61);
        gridPane.add(tAmt2A, 8, 61);
        gridPane.add(btCalculateQty2A, 3, 62);
        gridPane.add(btAmount2A, 6, 62);

        //12th horizontal separator
        final Separator sepHor12 = new Separator();
        sepHor12.setValignment(VPos.CENTER);
        GridPane.setConstraints(sepHor12, 0, 63);
        GridPane.setColumnSpan(sepHor12, 15);
        gridPane.getChildren().add(sepHor12);

        gridPane.add(snowCB, 0, 64);
        gridPane.add(new Label("Enter Expression:"), 0, 65);
        gridPane.add(expression11, 1, 65);
        gridPane.add(new Label("Quantity(sft):"), 2, 65);
        gridPane.add(result11, 3, 65);
        gridPane.add(new Label("Enter Rate % (it is per 100):"), 4, 65);
        gridPane.add(rate11, 5, 65);
        gridPane.add(new Label("Amount:"), 6, 65);
        gridPane.add(tAmt11, 7, 65);
        gridPane.add(btCalculateQty11, 1, 66);
        gridPane.add(btAmount11, 5, 66);

        //13th horizontal separator
        final Separator sepHor13 = new Separator();
        sepHor13.setValignment(VPos.CENTER);
        GridPane.setConstraints(sepHor13, 0, 67);
        GridPane.setColumnSpan(sepHor13, 15);
        gridPane.getChildren().add(sepHor13);

        gridPane.add(paintCB, 0, 68);
        gridPane.add(new Label("Enter Expression:"), 0, 69);
        gridPane.add(expression12, 1, 69);
        gridPane.add(new Label("Quantity(sft):"), 2, 69);
        gridPane.add(result12, 3, 69);
        gridPane.add(new Label("Enter Rate % (it is per 100):"), 4, 69);
        gridPane.add(rate12, 5, 69);
        gridPane.add(new Label("Amount:"), 6, 69);
        gridPane.add(tAmt12, 7, 69);
        gridPane.add(btCalculateQty12, 1, 70);
        gridPane.add(btAmount12, 5, 70);

        //14th horizontal separator
        final Separator sepHor14 = new Separator();
        sepHor14.setValignment(VPos.CENTER);
        GridPane.setConstraints(sepHor14, 0, 71);
        GridPane.setColumnSpan(sepHor14, 15);
        gridPane.getChildren().add(sepHor14);

        gridPane.add(graniteCB, 0, 72);
        gridPane.add(new Label("Please enter dimensions of granite"), 1, 72);
        gridPane.add(graniteSize, 2, 72);
        gridPane.add(new Label("Enter Expression:"), 0, 73);
        gridPane.add(expression13, 1, 73);
        gridPane.add(new Label("Quantity(cft):"), 2, 73);
        gridPane.add(result13, 3, 73);
        gridPane.add(new Label("Enter Rate % (it is per 100):"), 4, 73);
        gridPane.add(rate13, 5, 73);
        gridPane.add(new Label("Amount:"), 6, 73);
        gridPane.add(tAmt13, 7, 73);
        gridPane.add(btCalculateQty13, 1, 74);
        gridPane.add(btAmount13, 5, 74);

        //15th horizontal separator
        final Separator sepHor15 = new Separator();
        sepHor15.setValignment(VPos.CENTER);
        GridPane.setConstraints(sepHor15, 0, 75);
        GridPane.setColumnSpan(sepHor15, 15);
        gridPane.getChildren().add(sepHor15);
        gridPane.add(ultimateCal, 1, 76);
        gridPane.add(ultimateTarget, 2, 76);

        //16th horizontal separator
        final Separator sepHor16 = new Separator();
        sepHor16.setValignment(VPos.CENTER);
        GridPane.setConstraints(sepHor16, 0, 77);
        GridPane.setColumnSpan(sepHor16, 15);
        gridPane.getChildren().add(sepHor16);

        gridPane.add(xtraCB, 0, 78);
        gridPane.add(elecCB, 1, 79);
        gridPane.add(elecPerc, 2, 79);
        gridPane.add(sanitaryCB, 1, 80);
        gridPane.add(sanitaryPerc, 2, 80);
        gridPane.add(extraCB, 1, 81);
        gridPane.add(extraPerc, 2, 81);
        gridPane.add(new Label("Write Total Extra % Charges"), 1, 82);
        gridPane.add(totalPerc, 2, 82);
        gridPane.add(ultimateGrandVal, 1, 83);
        gridPane.add(displayUltimateGrandVal, 2, 83);

        //17th horizontal separator
        final Separator sepHor17 = new Separator();
        sepHor17.setValignment(VPos.CENTER);
        GridPane.setConstraints(sepHor17, 0, 84);
        GridPane.setColumnSpan(sepHor17, 15);
        gridPane.getChildren().add(sepHor17);
        gridPane.add(generateX, 0, 85);
        gridPane.add(generate, 1, 85);
        gridPane.add(actionTarget, 2, 85);

        // Set properties for UI
        gridPane.setAlignment(Pos.CENTER);
        Feet.setAlignment(Pos.BOTTOM_RIGHT);
        inches.setAlignment(Pos.BOTTOM_RIGHT);
        totalFeet.setAlignment(Pos.BOTTOM_RIGHT);
        btConvert.setAlignment(Pos.BOTTOM_RIGHT);

        totalFeet.setEditable(false);
        GridPane.setHalignment(btConvert, HPos.RIGHT);

        scrollPane.setHbarPolicy(ScrollPane.ScrollBarPolicy.ALWAYS);
        scrollPane.setVbarPolicy(ScrollPane.ScrollBarPolicy.ALWAYS);
        scrollPane.setContent(gridPane);

        // Process events
        btConvert.setOnAction(event -> Convert());

        btCalculateQty.setOnAction(event -> {
            result.setText(String.valueOf(new ExpressionBuilder(expression.getText()).build().evaluate()));
        });
        btAmount.setOnAction(event -> {
            tAmt.setText(String.valueOf((Double.valueOf(result.getText()) * Double.valueOf(rate.getText()) / 1000)));
            ultimate.add(Double.parseDouble(tAmt.getText()));
        });

        btCalculateQty1.setOnAction(event -> {
            result1.setText(String.valueOf(new ExpressionBuilder(expression1.getText()).build().evaluate()));
            prevRes.setText(String.valueOf(result1.getText()));
        });
        btAmount1.setOnAction(event -> {
            tAmt1.setText(String.valueOf((Double.valueOf(result1.getText()) * Double.valueOf(rate1.getText()) / 100)));
            ultimate.add(Double.parseDouble(tAmt1.getText()));
        });

        btCalculateQty2.setOnAction(event -> {
            result2.setText(String.valueOf(Double.valueOf(sandRate2.getText()) * Double.valueOf(prevRes.getText())));
        });
        btAmount2.setOnAction(event -> {
            tAmt2.setText(String.valueOf(Double.valueOf(result2.getText()) * Double.valueOf(rate2.getText()) / 100));
            ultimate.add(Double.parseDouble(tAmt2.getText()));
        });

        evalA.setOnAction(event -> {
            resA.setText(String.valueOf(new ExpressionBuilder(expressionA.getText()).build().evaluate()));
            a1.add(Double.parseDouble(resA.getText()));
        });

        evalB.setOnAction(event -> {
            resB.setText(String.valueOf(new ExpressionBuilder(expressionB.getText()).build().evaluate()));
            a1.add(Double.parseDouble(resB.getText()));
        });

        evalC.setOnAction(event -> {
            resC.setText(String.valueOf(new ExpressionBuilder(expressionC.getText()).build().evaluate()));
            a1.add(Double.parseDouble(resC.getText()));
        });

        evalD.setOnAction(event -> {
            resD.setText(String.valueOf(new ExpressionBuilder(expressionD.getText()).build().evaluate()));
            a1.add(Double.parseDouble(resD.getText()));
        });

        evalE.setOnAction(event -> {
            resE.setText(String.valueOf(new ExpressionBuilder(expressionE.getText()).build().evaluate()));
            a1.add(Double.parseDouble(resE.getText()));
        });

        evalF.setOnAction(event -> {
            resF.setText(String.valueOf(new ExpressionBuilder(expressionF.getText()).build().evaluate()));
            a1.add(Double.parseDouble(resF.getText()));
        });

        evalG.setOnAction(event -> {
            resG.setText(String.valueOf(new ExpressionBuilder(expressionG.getText()).build().evaluate()));
            a1.add(Double.parseDouble(resG.getText()));
        });

        evalH.setOnAction(event -> {
            resH.setText(String.valueOf(new ExpressionBuilder(expressionH.getText()).build().evaluate()));
            a1.add(Double.parseDouble(resH.getText()));
        });

        evalI.setOnAction(event -> {
            resI.setText(String.valueOf(new ExpressionBuilder(expressionI.getText()).build().evaluate()));
            a1.add(Double.parseDouble(resI.getText()));
        });

        evalJ.setOnAction(event -> {
            resJ.setText(String.valueOf(new ExpressionBuilder(expressionJ.getText()).build().evaluate()));
            a1.add(Double.parseDouble(resJ.getText()));
        });

        evalK.setOnAction(event -> {
            resK.setText(String.valueOf(new ExpressionBuilder(expressionK.getText()).build().evaluate()));
            a1.add(Double.parseDouble(resK.getText()));
        });

        btSumTotal.setOnAction(event -> {
            for (int i = 0; i < a1.size(); i++) {
                sum = sum + Double.parseDouble(a1.get(i).toString());
            }
            totalSum.setText(String.valueOf(sum));
            btSumTotal.setDisable(true);
        });

        btAmountSum.setOnAction(event -> {
            tAmtSum.setText(String.valueOf(sum * (Double.valueOf(sumRate.getText()) / 100)));
            ultimate.add(Double.parseDouble(tAmtSum.getText()));
        });

        //for 3rd one
        btCalculateQty3.setOnAction(event -> {
            result3.setText(String.valueOf(new ExpressionBuilder(expression3.getText()).build().evaluate()));
        });
        btAmount3.setOnAction(event -> {
            tAmt3.setText(String.valueOf((Double.valueOf(result3.getText()) * Double.valueOf(rate3.getText()) / 100)));
            ultimate.add(Double.parseDouble(tAmt3.getText()));
        });

        //for 4th one
        btCalculateQty4.setOnAction(event -> {
            result4.setText(String.valueOf(new ExpressionBuilder(expression4.getText()).build().evaluate()));
        });
        btAmount4.setOnAction(event -> {
            tAmt4.setText(String.valueOf((Double.valueOf(result4.getText()) * Double.valueOf(rate4.getText()) / 100)));
            ultimate.add(Double.parseDouble(tAmt4.getText()));
        });

        //for 5th one
        btCalculateQty5.setOnAction(event -> {
            result5.setText(String.valueOf(new ExpressionBuilder(expression5.getText()).build().evaluate()));
        });
        btAmount5.setOnAction(event -> {
            tAmt5.setText(String.valueOf((Double.valueOf(result5.getText()) * Double.valueOf(rate5.getText()) / 100)));
            ultimate.add(Double.parseDouble(tAmt5.getText()));
        });

        //for 6th one
        btCalculateQty6.setOnAction(event -> {
            result6.setText(String.valueOf(new ExpressionBuilder(expression6.getText()).build().evaluate()));
        });
        btAmount6.setOnAction(event -> {
            tAmt6.setText(String.valueOf((Double.valueOf(result6.getText()) * Double.valueOf(rate6.getText()) / 100)));
            ultimate.add(Double.parseDouble(tAmt6.getText()));
        });

        //for 7th one
        btCalculateQty7.setOnAction(event -> {
            result7.setText(String.valueOf(new ExpressionBuilder(expression7.getText()).build().evaluate()));
        });
        btAmount7.setOnAction(event -> {
            tAmt7.setText(String.valueOf((Double.valueOf(result7.getText()) * Double.valueOf(rate7.getText()))));
            ultimate.add(Double.parseDouble(tAmt7.getText()));
        });

        //for septic tank Lump sum cost
        btAmount8.setOnAction(event -> {
            tAmt8.setText(String.valueOf(String.valueOf(new ExpressionBuilder(expression8.getText()).build().evaluate())));
            ultimate.add(Double.parseDouble(tAmt8.getText()));
        });

        //for 9th one
        btCalculateQty9.setOnAction(event -> {
            result9.setText(String.valueOf(new ExpressionBuilder(expression9.getText()).build().evaluate()));
        });
        btAmount9.setOnAction(event -> {
            tAmt9.setText(String.valueOf((Double.valueOf(result9.getText()) * Double.valueOf(rate9.getText()))));
            ultimate.add(Double.parseDouble(tAmt9.getText()));
        });

        //for 10th one
        btCalculateQty10.setOnAction(event -> {
            result10.setText(String.valueOf(new ExpressionBuilder(expression10.getText()).build().evaluate()));
            prevResA.setText(String.valueOf(result10.getText()));
        });
        btAmount10.setOnAction(event -> {
            tAmt10.setText(String.valueOf((Double.valueOf(result10.getText()) * Double.valueOf(rate10.getText()) / 100)));
            ultimate.add(Double.parseDouble(tAmt10.getText()));
        });
        // for 2nd part in ips calculation
        btCalculateQty2A.setOnAction(event -> {
            result2A.setText(String.valueOf(Double.valueOf(sandRate2A.getText()) * Double.valueOf(prevResA.getText())));
        });
        btAmount2A.setOnAction(event -> {
            tAmt2A.setText(String.valueOf(Double.valueOf(result2A.getText()) * Double.valueOf(rate2A.getText()) / 100));
            ultimate.add(Double.parseDouble(tAmt2A.getText()));
        });

        //for 11th one
        btCalculateQty11.setOnAction(event -> {
            result11.setText(String.valueOf(new ExpressionBuilder(expression11.getText()).build().evaluate()));
        });
        btAmount11.setOnAction(event -> {
            tAmt11.setText(String.valueOf((Double.valueOf(result11.getText()) * Double.valueOf(rate11.getText()) / 100)));
            ultimate.add(Double.parseDouble(tAmt11.getText()));
        });

        //for 12th one
        btCalculateQty12.setOnAction(event -> {
            result12.setText(String.valueOf(new ExpressionBuilder(expression12.getText()).build().evaluate()));
        });
        btAmount12.setOnAction(event -> {
            tAmt12.setText(String.valueOf((Double.valueOf(result12.getText()) * Double.valueOf(rate12.getText()) / 100)));
            ultimate.add(Double.parseDouble(tAmt12.getText()));
        });

        //for 13th one
        btCalculateQty13.setOnAction(event -> {
            result13.setText(String.valueOf(new ExpressionBuilder(expression13.getText()).build().evaluate()));
        });
        btAmount13.setOnAction(event -> {
            tAmt13.setText(String.valueOf((Double.valueOf(result13.getText()) * Double.valueOf(rate13.getText()) / 100)));
            ultimate.add(Double.parseDouble(tAmt13.getText()));
        });

        //for cost calculation before adding extras
        ultimateCal.setOnAction(event -> {
            for (int i = 0; i < ultimate.size(); i++) {
                ultimateSum = ultimateSum + Double.parseDouble(ultimate.get(i).toString());
            }
            ultimateTarget.setText(String.valueOf(ultimateSum));
            ultimateCal.setDisable(true);
        });

        //for calculation after adding extra charges
        ultimateGrandVal.setOnAction(event -> {
            ultimateGrandSum = ultimateSum + (Double.valueOf(totalPerc.getText()) / 100 * ultimateSum);
            displayUltimateGrandVal.setText(String.valueOf(ultimateGrandSum));
        });

        //for generating the document
        generate.setOnAction(event -> {
            try {
                Generate();
            } catch (IOException e) {
                e.printStackTrace();
            }
        });

        generateX.setOnAction(event -> {
            try {
                GenerateX();
            } catch (IOException e) {
                e.printStackTrace();
            }
        });

        ewCB.setOnAction(new EventHandler<ActionEvent>() {
            @Override
            public void handle(ActionEvent event) {
                ewFlag = true;
            }
        });

        bfCB.setOnAction(new EventHandler<ActionEvent>() {
            @Override
            public void handle(ActionEvent event) {
                bfFlag = true;
            }
        });

        rcCB.setOnAction(new EventHandler<ActionEvent>() {
            @Override
            public void handle(ActionEvent event) {
                rcFlag = true;
            }
        });

        thckBrkCB.setOnAction(new EventHandler<ActionEvent>() {
            @Override
            public void handle(ActionEvent event) {
                thckBrkFlag = true;
            }
        });

        thnBrkCB.setOnAction(new EventHandler<ActionEvent>() {
            @Override
            public void handle(ActionEvent event) {
                thnBrkFlag = true;
            }
        });

        halfCB.setOnAction(new EventHandler<ActionEvent>() {
            @Override
            public void handle(ActionEvent event) {
                halfFlag = true;
            }
        });

        qtrCB.setOnAction(new EventHandler<ActionEvent>() {
            @Override
            public void handle(ActionEvent event) {
                qtrFlag = true;
            }
        });

        steelCB.setOnAction(new EventHandler<ActionEvent>() {
            @Override
            public void handle(ActionEvent event) {
                steelFlag = true;
            }
        });

        tankCB.setOnAction(new EventHandler<ActionEvent>() {
            @Override
            public void handle(ActionEvent event) {
                tankFlag = true;
            }
        });

        woodCB.setOnAction(new EventHandler<ActionEvent>() {
            @Override
            public void handle(ActionEvent event) {
                woodFlag = true;
            }
        });

        ipsCB.setOnAction(new EventHandler<ActionEvent>() {
            @Override
            public void handle(ActionEvent event) {
                ipsFlag = true;
            }
        });

        tankCB.setOnAction(new EventHandler<ActionEvent>() {
            @Override
            public void handle(ActionEvent event) {
                tankFlag = true;
            }
        });
        xtraCB.setOnAction(new EventHandler<ActionEvent>() {
            @Override
            public void handle(ActionEvent event) {
                xtraFlag = true;
            }
        });

        snowCB.setOnAction(new EventHandler<ActionEvent>() {
            @Override
            public void handle(ActionEvent event) {
                snowFlag = true;
            }
        });

        paintCB.setOnAction(new EventHandler<ActionEvent>() {
            @Override
            public void handle(ActionEvent event) {
                paintFlag = true;
            }
        });

        graniteCB.setOnAction(new EventHandler<ActionEvent>() {
            @Override
            public void handle(ActionEvent event) {
                graniteFlag = true;
            }
        });

        //for sub menu part of R.C.C.
        aCB.setOnAction(new EventHandler<ActionEvent>() {
            @Override
            public void handle(ActionEvent event) {
                aFlag = true;
            }
        });

        bCB.setOnAction(new EventHandler<ActionEvent>() {
            @Override
            public void handle(ActionEvent event) {
                bFlag = true;
            }
        });

        cCB.setOnAction(new EventHandler<ActionEvent>() {
            @Override
            public void handle(ActionEvent event) {
                cFlag = true;
            }
        });

        dCB.setOnAction(new EventHandler<ActionEvent>() {
            @Override
            public void handle(ActionEvent event) {
                dFlag = true;
            }
        });

        eCB.setOnAction(new EventHandler<ActionEvent>() {
            @Override
            public void handle(ActionEvent event) {
                eFlag = true;
            }
        });

        fCB.setOnAction(new EventHandler<ActionEvent>() {
            @Override
            public void handle(ActionEvent event) {
                fFlag = true;
            }
        });

        gCB.setOnAction(new EventHandler<ActionEvent>() {
            @Override
            public void handle(ActionEvent event) {
                gFlag = true;
            }
        });

        hCB.setOnAction(new EventHandler<ActionEvent>() {
            @Override
            public void handle(ActionEvent event) {
                hFlag = true;
            }
        });

        iCB.setOnAction(new EventHandler<ActionEvent>() {
            @Override
            public void handle(ActionEvent event) {
                iFlag = true;
            }
        });

        jCB.setOnAction(new EventHandler<ActionEvent>() {
            @Override
            public void handle(ActionEvent event) {
                jFlag = true;
            }
        });

        kCB.setOnAction(new EventHandler<ActionEvent>() {
            @Override
            public void handle(ActionEvent event) {
                kFlag = true;
            }
        });

        //for sub menu part of extra charges
        elecCB.setOnAction(new EventHandler<ActionEvent>() {
            @Override
            public void handle(ActionEvent event) {
                elecFlag = true;
            }
        });

        sanitaryCB.setOnAction(new EventHandler<ActionEvent>() {
            @Override
            public void handle(ActionEvent event) {
                sanitaryFlag = true;
            }
        });

        extraCB.setOnAction(new EventHandler<ActionEvent>() {
            @Override
            public void handle(ActionEvent event) {
                extraFlag = true;
            }
        });

        // Create a scene and place it in the stage
        Scene scene = new Scene(scrollPane, 1024, 768);
        primaryStage.getIcons().add(new Image("/home.png")); //set icon for software
        primaryStage.setTitle("Er. R.K. Sinha Home Estimation Software"); // Set title
        primaryStage.setScene(scene); // Place the scene in the stage
        primaryStage.show(); // Display the stage
    }

    private void Convert() {
        // Get values from text fields
        double interest = Double.parseDouble(Feet.getText());
        double FeetAmount = Double.parseDouble(inches.getText());
        double result = ((interest * 12) + FeetAmount) / 12;
        // Display the result
        totalFeet.setText(String.format("%.2f", result));
    }

    private void Generate() throws IOException {
        actionTarget.setText("Word Document generated");
        showSaveFileChooser();
    }

    private void GenerateX() throws IOException {
        actionTarget.setText("Excel Sheet generated");
        showSaveFileChooserX();
    }

    private void showSaveFileChooser() throws IOException {
        FileChooser fileChooser = new FileChooser();
        fileChooser.setTitle("Save file As...");
        fileChooser.setInitialFileName(defaultFileName);
        FileChooser.ExtensionFilter extFilter = new FileChooser.ExtensionFilter("Word file (*.docx)", "*.docx");
        fileChooser.getExtensionFilters().add(extFilter);
        File savedFile = fileChooser.showSaveDialog(savedStage);

        //calling myFileWriter method which contains the logic for creating paragraph
        XWPFDocument document1 = myFileWriter();

        if (savedFile != null) {

            try {
                //Write the Document in file system
                FileOutputStream out = new FileOutputStream(new File(savedFile.toString()));
                document1.write(out);
                out.close();
                //saveFileRoutine(savedFile);
            } catch (IOException e) {
                e.printStackTrace();
                actionTarget.setText("An ERROR occurred while saving the file!" + savedFile.toString());
                return;
            }
            actionTarget.setText("File saved: " + savedFile.toString());
        } else {
            actionTarget.setText("File save cancelled.");
        }
    }

    private void showSaveFileChooserX() throws IOException {
        FileChooser fileChooser = new FileChooser();
        fileChooser.setTitle("Save file As...");
        fileChooser.setInitialFileName(defaultFileName);
        FileChooser.ExtensionFilter extFilter = new FileChooser.ExtensionFilter("Excel file (*.xlsx)", "*.xlsx");
        fileChooser.getExtensionFilters().add(extFilter);
        File savedFile = fileChooser.showSaveDialog(savedStage);

        //calling myFileWriter method which contains the logic for creating paragraph
        XSSFWorkbook document1 = myFileWriterX();

        if (savedFile != null) {

            try {
                //Write the Document in file system
                FileOutputStream out = new FileOutputStream(new File(savedFile.toString()));
                document1.write(out);
                out.close();
                //saveFileRoutine(savedFile);
            } catch (IOException e) {
                e.printStackTrace();
                actionTarget.setText("An ERROR occurred while saving the file!" + savedFile.toString());
                return;
            }
            actionTarget.setText("File saved: " + savedFile.toString());
        } else {
            actionTarget.setText("File save cancelled.");
        }
    }

    //For excel version of estimate

    private XSSFWorkbook myFileWriterX() throws IOException {
        //Blank workbook
        XSSFWorkbook workbook = new XSSFWorkbook();

        //create sheet
        XSSFSheet spreadsheet = workbook.createSheet("Estimate");

        //Create row object
        XSSFRow row;

        //set print area with indexes
//        workbook.setPrintArea(
//                0, //sheet index
//                0, //start column
//                4, //end column
//                0, //start row
//                100 //end row
//        );

        //set paper size
        spreadsheet.getPrintSetup().setPaperSize(XSSFPrintSetup.A4_PAPERSIZE);
        //set display grid lines or not
        spreadsheet.setDisplayGridlines(true);
        //set print grid lines or not
        spreadsheet.setPrintGridlines(true);

        //MERGING CELLS
        spreadsheet.addMergedRegion(new CellRangeAddress(
                0, //first row (0-based)
                0, //last row (0-based)
                0, //first column (0-based)
                5 //last column (0-based)
        ));
        spreadsheet.addMergedRegion(new CellRangeAddress(
                1, //first row (0-based)
                1, //last row (0-based)
                0, //first column (0-based)
                5 //last column (0-based)
        ));
        spreadsheet.addMergedRegion(new CellRangeAddress(
                2, //first row (0-based)
                2, //last row (0-based)
                0, //first column (0-based)
                5 //last column (0-based)
        ));
        spreadsheet.addMergedRegion(new CellRangeAddress(
                3, //first row (0-based)
                3, //last row (0-based)
                0, //first column (0-based)
                5 //last column (0-based)
        ));

        CellStyle style = workbook.createCellStyle();//Create style
        Font font = workbook.createFont();//Create font
        font.setBold(true);//Make font bold
        style.setFont(font);//set it to bold

        row = spreadsheet.createRow(0);
        row.createCell((short) 0).setCellValue("Estimated cost of " + custName.getText() + " " + custDetails.getText());

        row = spreadsheet.createRow(1);
        row.createCell((short) 0).setCellValue("BUILT UP G.F AREA- " + builtUp.getText() + " SFT");

        row = spreadsheet.createRow(2);
        row.createCell((short) 0).setCellValue("Abstract of cost (Based on current market rates)");

        row = spreadsheet.createRow(4);
        row.createCell((short) 0).setCellValue("Sl.No");
        row.createCell((short) 1).setCellValue("Description");
        row.createCell((short) 2).setCellValue("Quantity");
        row.createCell((short) 3).setCellValue("Rate");
        row.createCell((short) 4).setCellValue("Unit");
        row.createCell((short) 5).setCellValue("Amount");

        for (int i = 0; i <= 5; i++) {//For each cell in the row
            row.getCell(i).setCellStyle(style);//Set the style
        }

        if (ewFlag) {
            row = spreadsheet.createRow(rowIndex);
            row.createCell((short) 0).setCellValue((indexX + "."));
            rowIndex++;
            indexX++;
            row.createCell((short) 1).setCellValue("E/W in excavation of all types of" +
                    " soil for both Main & support pocket" +
                    " foundation including all cost of" +
                    " labour & materials, etc. completion job");
            //row.setHeight((short) 8000);
            row = spreadsheet.createRow(rowIndex);
            row.createCell((short) 1).setCellValue(expression.getText().replace("*", "X"));
            row.createCell((short) 2).setCellValue(result.getText() + "cft");
            row.createCell((short) 3).setCellValue(rate.getText());
            row.createCell((short) 4).setCellValue("%cft");
            row.createCell((short) 5).setCellValue(tAmt.getText());
        }

        if (bfFlag) {
            rowIndex = rowIndex + 2;
            row = spreadsheet.createRow(rowIndex);
            row.createCell((short) 0).setCellValue((indexX + "."));
            rowIndex++;
            indexX++;
            row.createCell((short) 1).setCellValue("Supplying & providing 1st class B.F." +
                    " Soling with joints filled with local sand" +
                    " on 6” Local sand including all cost of" +
                    " labour & materials, etc. comp job");
            row = spreadsheet.createRow(rowIndex);
            row.createCell((short) 1).setCellValue(expression1.getText().replace("*", "X"));
            row.createCell((short) 2).setCellValue(result1.getText() + "cft");
            row.createCell((short) 3).setCellValue(rate1.getText());
            row.createCell((short) 4).setCellValue("%cft");
            row.createCell((short) 5).setCellValue(tAmt1.getText());
            rowIndex++;

            row = spreadsheet.createRow(rowIndex);
            row.createCell((short) 1).setCellValue("Local sand: " + result1.getText() + "X" + sandRate2.getText());
            row.createCell((short) 2).setCellValue(result2.getText() + "cft");
            row.createCell((short) 3).setCellValue(rate2.getText());
            row.createCell((short) 4).setCellValue("%cft");
            row.createCell((short) 5).setCellValue(tAmt2.getText());
        }

        if (rcFlag) {
            rowIndex = rowIndex + 2;
            row = spreadsheet.createRow(rowIndex);
            row.createCell((short) 0).setCellValue((indexX + "."));
            rowIndex++;
            indexX++;
            row.createCell((short) 1).setCellValue("Supplying & providing R.C.C(1:2:4) with stone" +
                    " chips, sone-sand including all cost of labour &" +
                    " materials curing, shuttering etc. but excluding the" +
                    " cost of reinforcement of following items:");
            if (aFlag) {
                row = spreadsheet.createRow(rowIndex);
                row.createCell((short) 1).setCellValue(subIndex + ". Pocket (main & support)" + " = " + expressionA.getText().replace("*", "X") + " = " + resA.getText() + "cft" + "    ");
                subIndexX++;
                rowIndex++;
            }

            if (bFlag) {
                row = spreadsheet.createRow(rowIndex);
                row.createCell((short) 1).setCellValue(subIndex + ". 10”x12” G.B." + " = " + expressionB.getText().replace("*", "X") + " = " + resB.getText() + "cft" + "    ");
                subIndexX++;
                rowIndex++;
            }

            if (cFlag) {
                row = spreadsheet.createRow(rowIndex);
                row.createCell((short) 1).setCellValue(subIndex + ". 10”x10” Stiffeners" + " = " + expressionC.getText().replace("*", "X") + " = " + resC.getText() + "cft" + "    ");
                subIndexX++;
                rowIndex++;
            }

            if (dFlag) {
                row = spreadsheet.createRow(rowIndex);
                row.createCell((short) 1).setCellValue(subIndex + ". 10”x6” Lintel band" + " = " + expressionD.getText().replace("*", "X") + " = " + resD.getText() + "cft" + "    ");
                subIndexX++;
                rowIndex++;
            }

            if (eFlag) {
                row = spreadsheet.createRow(rowIndex);
                row.createCell((short) 1).setCellValue(subIndex + ". Chajja, loft, etc." + " = " + expressionE.getText().replace("*", "X") + " = " + resE.getText() + "cft" + "    ");
                subIndexX++;
                rowIndex++;
            }

            if (fFlag) {
                row = spreadsheet.createRow(rowIndex);
                row.createCell((short) 1).setCellValue(subIndex + ". Roof slab" + " = " + expressionF.getText().replace("*", "X") + " = " + resF.getText() + "cft" + "    ");
                subIndexX++;
                rowIndex++;
            }

            if (gFlag) {
                row = spreadsheet.createRow(rowIndex);
                row.createCell((short) 1).setCellValue(subIndex + ". Stair case" + " = " + expressionG.getText().replace("*", "X") + " = " + resG.getText() + "cft" + "    ");
                subIndexX++;
                rowIndex++;
            }

            if (hFlag) {
                row = spreadsheet.createRow(rowIndex);
                row.createCell((short) 1).setCellValue(subIndex + ". Front Beam" + " = " + expressionH.getText().replace("*", "X") + " = " + resH.getText() + "cft" + "    ");
                subIndexX++;
                rowIndex++;
            }

            if (iFlag) {
                row = spreadsheet.createRow(rowIndex);
                row.createCell((short) 1).setCellValue(subIndex + ". 10”x10” beam on 5” B/W under roof slab" + " = " + expressionI.getText().replace("*", "X") + " = " + resI.getText() + "cft" + "    ");
                subIndexX++;
                rowIndex++;
            }

            if (jFlag) {
                row = spreadsheet.createRow(rowIndex);
                row.createCell((short) 1).setCellValue(subIndex + ". Veranda beam " + " = " + expressionJ.getText().replace("*", "X") + " = " + resJ.getText() + "cft" + "    ");
                subIndexX++;
                rowIndex++;
            }

            if (kFlag) {
                row = spreadsheet.createRow(rowIndex);
                row.createCell((short) 1).setCellValue(subIndex + ". R.C.C. railing through roof" + " = " + expressionK.getText().replace("*", "X") + " = " + resK.getText() + "cft" + "    ");
                subIndexX++;
                rowIndex++;
            }

            row = spreadsheet.createRow(rowIndex);
            row.createCell((short) 1).setCellValue("Total= " + totalSum.getText() + "cft");
            row.createCell((short) 2).setCellValue(totalSum.getText() + "cft");
            row.createCell((short) 3).setCellValue(sumRate.getText());
            row.createCell((short) 4).setCellValue("%cft");
            row.createCell((short) 5).setCellValue(tAmtSum.getText());
        }

        if (thckBrkFlag) {
            rowIndex = rowIndex + 2;
            row = spreadsheet.createRow(rowIndex);
            row.createCell((short) 0).setCellValue((indexX + "."));
            rowIndex++;
            indexX++;
            row.createCell((short) 1).setCellValue("Supplying & providing 10” brick work in C.M. (1:6)" +
                    " with sone sand and 1st class bricks including all" +
                    " cost of labour & materials, curing etc. completion job");

            row = spreadsheet.createRow(rowIndex);
            row.createCell((short) 1).setCellValue(expression3.getText().replace("*", "X"));
            row.createCell((short) 2).setCellValue(result3.getText() + "cft");
            row.createCell((short) 3).setCellValue(rate3.getText());
            row.createCell((short) 4).setCellValue("%cft");
            row.createCell((short) 5).setCellValue(tAmt3.getText());
        }

        if (thnBrkFlag) {
            rowIndex = rowIndex + 2;
            row = spreadsheet.createRow(rowIndex);
            row.createCell((short) 0).setCellValue((indexX + "."));
            rowIndex++;
            indexX++;
            row.createCell((short) 1).setCellValue("Supplying & providing 5” B/W in C.M. (1:5)" +
                    " with sone-sand and 1st class bricks including all" +
                    " cost of labour & materials, curing etc. completion job");

            row = spreadsheet.createRow(rowIndex);
            row.createCell((short) 1).setCellValue(expression4.getText().replace("*", "X"));
            row.createCell((short) 2).setCellValue(result4.getText() + "cft");
            row.createCell((short) 3).setCellValue(rate4.getText());
            row.createCell((short) 4).setCellValue("%cft");
            row.createCell((short) 5).setCellValue(tAmt4.getText());
        }

        if (halfFlag) {
            rowIndex = rowIndex + 2;
            row = spreadsheet.createRow(rowIndex);
            row.createCell((short) 0).setCellValue((indexX + "."));
            rowIndex++;
            indexX++;
            row.createCell((short) 1).setCellValue("Supplying & providing 1\\2”C.P.in C.M. (1:6) with" +
                    " sone-sand both sides including all cost of" +
                    " labour & materials for both sides");

            row = spreadsheet.createRow(rowIndex);
            row.createCell((short) 1).setCellValue(expression5.getText().replace("*", "X"));
            row.createCell((short) 2).setCellValue(result5.getText() + "cft");
            row.createCell((short) 3).setCellValue(rate5.getText());
            row.createCell((short) 4).setCellValue("%cft");
            row.createCell((short) 5).setCellValue(tAmt5.getText());
        }

        if (qtrFlag) {
            rowIndex = rowIndex + 2;
            row = spreadsheet.createRow(rowIndex);
            row.createCell((short) 0).setCellValue((indexX + "."));
            rowIndex++;
            indexX++;
            row.createCell((short) 1).setCellValue("Supplying & providing 1\\4” C.P in C.M. (1:4) with" +
                    " sone-sand including all cost of labour & materials," +
                    " curing completion job.");

            row = spreadsheet.createRow(rowIndex);
            row.createCell((short) 1).setCellValue(expression6.getText().replace("*", "X"));
            row.createCell((short) 2).setCellValue(result6.getText() + "cft");
            row.createCell((short) 3).setCellValue(rate6.getText());
            row.createCell((short) 4).setCellValue("%cft");
            row.createCell((short) 5).setCellValue(tAmt6.getText());
        }

        if (steelFlag) {
            rowIndex = rowIndex + 2;
            row = spreadsheet.createRow(rowIndex);
            row.createCell((short) 0).setCellValue((indexX + "."));
            rowIndex++;
            indexX++;
            row.createCell((short) 1).setCellValue("Supplying & providing different diameter of steel" +
                    " to the site including all cost of labour & materials," +
                    " transportation cost, taxes etc.");

            row = spreadsheet.createRow(rowIndex);
            row.createCell((short) 1).setCellValue(expression7.getText().replace("*", "X") + "M.T.");
            row.createCell((short) 2).setCellValue(result7.getText() + "M.T.");
            row.createCell((short) 3).setCellValue(rate7.getText());
            row.createCell((short) 4).setCellValue("P.M.T.");
            row.createCell((short) 5).setCellValue(tAmt7.getText());
        }

        if (tankFlag) {
            rowIndex = rowIndex + 2;
            row = spreadsheet.createRow(rowIndex);
            row.createCell((short) 0).setCellValue((indexX + "."));
            rowIndex++;
            indexX++;
            row.createCell((short) 1).setCellValue("Provision of " + tankSize.getText() + " septic tank" +
                    " with inlet outlet including all cost of" +
                    " labour & materials completion Job.");

            row = spreadsheet.createRow(rowIndex);
            row.createCell((short) 1).setCellValue(" Lump Sum cost = " + tAmt8.getText());
            row.createCell((short) 5).setCellValue(tAmt8.getText());
        }

        if (woodFlag) {
            rowIndex = rowIndex + 2;
            row = spreadsheet.createRow(rowIndex);
            row.createCell((short) 0).setCellValue((indexX + "."));
            rowIndex++;
            indexX++;
            row.createCell((short) 1).setCellValue("Supplying & providing with fixing well seasoned" +
                    " sal wood for chaukhat & panel for doors & windows" +
                    " including all cost of labour & materials completion job");

            row = spreadsheet.createRow(rowIndex);
            row.createCell((short) 1).setCellValue(expression9.getText().replace("*", "X"));
            row.createCell((short) 2).setCellValue(result9.getText() + "cft");
            row.createCell((short) 3).setCellValue(rate9.getText());
            row.createCell((short) 4).setCellValue("%cft");
            row.createCell((short) 5).setCellValue(tAmt9.getText());
        }

        if (ipsFlag) {
            rowIndex = rowIndex + 2;
            row = spreadsheet.createRow(rowIndex);
            row.createCell((short) 0).setCellValue((indexX + "."));
            rowIndex++;
            indexX++;
            row.createCell((short) 1).setCellValue("Supplying & Providing 1” I.P.S(1:2:4) with punning" +
                    " On 3” P.C.C.(1:3:6) with sone-sand,chips etc. including" +
                    " all cost of labour & materials completion job");

            row = spreadsheet.createRow(rowIndex);
            row.createCell((short) 1).setCellValue(expression10.getText().replace("*", "X"));
            row.createCell((short) 2).setCellValue(result10.getText() + "cft");
            row.createCell((short) 3).setCellValue(rate10.getText());
            row.createCell((short) 4).setCellValue("%cft");
            row.createCell((short) 5).setCellValue(tAmt10.getText());
            rowIndex++;

            row = spreadsheet.createRow(rowIndex);
            row.createCell((short) 1).setCellValue("3” P.C.C.(1:3:6)" + result10.getText() + "X" + sandRate2A.getText());
            row.createCell((short) 2).setCellValue(result2A.getText() + "cft");
            row.createCell((short) 3).setCellValue(rate2A.getText());
            row.createCell((short) 4).setCellValue("%cft");
            row.createCell((short) 5).setCellValue(tAmt2A.getText());
        }

        if (snowFlag) {
            rowIndex = rowIndex + 2;
            row = spreadsheet.createRow(rowIndex);
            row.createCell((short) 0).setCellValue((indexX + "."));
            rowIndex++;
            indexX++;
            row.createCell((short) 1).setCellValue("Two coats of snow cement to walls & ceiling" +
                    " including all cost of labour & materials completion Job");

            row = spreadsheet.createRow(rowIndex);
            row.createCell((short) 1).setCellValue(expression11.getText().replace("*", "X"));
            row.createCell((short) 2).setCellValue(result11.getText() + "cft");
            row.createCell((short) 3).setCellValue(rate11.getText());
            row.createCell((short) 4).setCellValue("%cft");
            row.createCell((short) 5).setCellValue(tAmt11.getText());
        }

        if (paintFlag) {
            rowIndex = rowIndex + 2;
            row = spreadsheet.createRow(rowIndex);
            row.createCell((short) 0).setCellValue((indexX + "."));
            rowIndex++;
            indexX++;
            row.createCell((short) 1).setCellValue("Supplying & Providing two coats of enamel paint to chaukhats & pannels" +
                    " including all cost of labour & materials completion Job");

            row = spreadsheet.createRow(rowIndex);
            row.createCell((short) 1).setCellValue(expression12.getText().replace("*", "X"));
            row.createCell((short) 2).setCellValue(result12.getText() + "cft");
            row.createCell((short) 3).setCellValue(rate12.getText());
            row.createCell((short) 4).setCellValue("%cft");
            row.createCell((short) 5).setCellValue(tAmt12.getText());
        }

        if (graniteFlag) {
            rowIndex = rowIndex + 2;
            row = spreadsheet.createRow(rowIndex);
            row.createCell((short) 0).setCellValue((indexX + "."));
            rowIndex++;
            indexX++;
            row.createCell((short) 1).setCellValue("Supplying & Providing " + graniteSize.getText() + " Granite on 3” P.C.C.\n" +
                    " (1:3:6) with sone-sand, chips etc including all\n" +
                    " cost of labors & materials completion Job");

            row = spreadsheet.createRow(rowIndex);
            row.createCell((short) 1).setCellValue(expression13.getText().replace("*", "X"));
            row.createCell((short) 2).setCellValue(result13.getText() + "cft");
            row.createCell((short) 3).setCellValue(rate13.getText());
            row.createCell((short) 4).setCellValue("%cft");
            row.createCell((short) 5).setCellValue(tAmt13.getText());
        }

        rowIndex = rowIndex + 2;
        row = spreadsheet.createRow(rowIndex);
        row.createCell((short) 1).setCellValue("TOTAL = ");
        row.createCell((short) 5).setCellValue(String.valueOf(ultimateSum));

        if (xtraFlag) {
            rowIndex = rowIndex + 2;
            row = spreadsheet.createRow(rowIndex);
            row.createCell((short) 1).setCellValue("EXTRA SERVICE CHARGES");

            if (elecFlag) {
                rowIndex++;
                row = spreadsheet.createRow(rowIndex);
                row.createCell((short) 1).setCellValue("Add " + elecPerc.getText() + " % for electrical services");
            }
            if (sanitaryFlag) {
                rowIndex++;
                row = spreadsheet.createRow(rowIndex);
                row.createCell((short) 1).setCellValue("Add " + sanitaryPerc.getText() + " % for sanitary services");
            }
            if (extraFlag) {
                rowIndex++;
                row = spreadsheet.createRow(rowIndex);
                row.createCell((short) 1).setCellValue("Add " + extraPerc.getText() + " % for extra services");
            }
            rowIndex++;
            row = spreadsheet.createRow(rowIndex);
            row.createCell((short) 1).setCellValue("Total " + totalPerc.getText() + " % charges extra");
        }
        rowIndex = rowIndex + 2;
        row = spreadsheet.createRow(rowIndex);
        row.createCell((short) 1).setCellValue("GRAND TOTAL = ");
        row.createCell((short) 5).setCellValue(String.valueOf(ultimateGrandSum));

        // for auto-sizing the column
        for (int i = 0; i < 100; i++) {
            spreadsheet.autoSizeColumn(i);
        }
        return workbook;
    }

    //for word version of file
    private XWPFDocument myFileWriter() throws IOException {
        //Blank Document
        XWPFDocument document = new XWPFDocument();

        //create paragraph
        XWPFParagraph paragraph = document.createParagraph();

        //Set alignment paragraph to LEFT
        paragraph.setAlignment(ParagraphAlignment.LEFT);
        XWPFRun run = paragraph.createRun();
        run.setBold(true);
        run.setFontSize(24);
        run.setText("Estimated cost of " + custName.getText() + " " + custDetails.getText());

        //Create Another paragraph
        paragraph = document.createParagraph();
        //Set alignment paragraph to RIGHT
        paragraph.setAlignment(ParagraphAlignment.CENTER);
        run = paragraph.createRun();
        run.setBold(true);
        run.setText(" BUILT UP G.F AREA- " + builtUp.getText() + " SFT");

        //Create Another paragraph
        paragraph = document.createParagraph();
        //Set alignment paragraph to RIGHT
        paragraph.setAlignment(ParagraphAlignment.CENTER);
        run = paragraph.createRun();
        run.setText("ABSTRACT OF COST (Based on current market rates)");

        if (ewFlag) {
            //1 Create Another paragraph
            paragraph = document.createParagraph();
            //Set alignment paragraph to RIGHT
            paragraph.setAlignment(ParagraphAlignment.LEFT);
            run = paragraph.createRun();
            run.setText(index + ". E/W in excavation of all types of" +
                    " soil for both Main & support pocket" +
                    " foundation including all cost of" +
                    " labour & materials, etc. completion job");
            index++;
            paragraph = document.createParagraph();
            run = paragraph.createRun();
            run.setText(expression.getText().replace("*", "X") + "    ");
            run.setText(" =" + result.getText() + "cft" + "    ");
            run.setText(rate.getText() + "%cft" + "    ");
            run.setText(tAmt.getText());
        }

        if (bfFlag) {
            //2 Create Another paragraph
            paragraph = document.createParagraph();
            //Set alignment paragraph to RIGHT
            paragraph.setAlignment(ParagraphAlignment.LEFT);
            run = paragraph.createRun();
            run.setText(index + ". Supplying & providing 1st class B.F." +
                    " Soling with joints filled with local sand" +
                    " on 6” Local sand including all cost of" +
                    " labour & materials, etc. comp job");
            index++;
            paragraph = document.createParagraph();
            run = paragraph.createRun();
            run.setText(expression1.getText().replace("*", "X") + "    ");
            run.setText(" =" + result1.getText() + "sft" + "    ");
            run.setText(rate1.getText() + "%sft" + "    ");
            run.setText(tAmt1.getText());

            paragraph = document.createParagraph();
            run = paragraph.createRun();
            run.setText("Local sand ");
            run.setText(result1.getText() + "x");
            run.setText(sandRate2.getText() + " =    ");
            run.setText(" =" + result2.getText() + "cft" + "   ");
            run.setText(rate2.getText() + "%cft" + "   ");
            run.setText(tAmt2.getText());
        }

        if (rcFlag) {
            //3 Create Another paragraph
            paragraph = document.createParagraph();
            //Set alignment paragraph to RIGHT
            paragraph.setAlignment(ParagraphAlignment.LEFT);
            run = paragraph.createRun();
            run.setText(index + ". Supplying & providing R.C.C(1:2:4) with stone" +
                    " chips, sone-sand including all cost of labour &" +
                    " materials curing, shuttering etc. but excluding the" +
                    " cost of reinforcement of following items:");
            index++;

            if (aFlag) {
                //Create Another paragraph
                paragraph = document.createParagraph();
                //Set alignment paragraph to RIGHT
                paragraph.setAlignment(ParagraphAlignment.LEFT);
                run = paragraph.createRun();
                run.setText(subIndex + ". Pocket (main & support)");
                subIndex++;

                run.setText(" = " + expressionA.getText().replace("*", "X") + "cft" + "    ");
                run.setText(" = " + resA.getText() + "cft" + "    ");
            }

            if (bFlag) {
                //Create Another paragraph
                paragraph = document.createParagraph();
                //Set alignment paragraph to RIGHT
                paragraph.setAlignment(ParagraphAlignment.LEFT);
                run = paragraph.createRun();
                run.setText(subIndex + ". 10”x12” G.B.");
                subIndex++;

                run.setText(" = " + expressionB.getText().replace("*", "X") + "    ");
                run.setText(" = " + resB.getText() + "cft" + "    ");
            }

            if (cFlag) {
                //Create Another paragraph
                paragraph = document.createParagraph();
                //Set alignment paragraph to RIGHT
                paragraph.setAlignment(ParagraphAlignment.LEFT);
                run = paragraph.createRun();
                run.setText(subIndex + ". 10”x10” Stiffeners");
                subIndex++;

                run.setText(" = " + expressionC.getText().replace("*", "X") + "    ");
                run.setText(" = " + resC.getText() + "cft" + "    ");
            }

            if (dFlag) {
                //Create Another paragraph
                paragraph = document.createParagraph();
                //Set alignment paragraph to RIGHT
                paragraph.setAlignment(ParagraphAlignment.LEFT);
                run = paragraph.createRun();
                run.setText(subIndex + ". 10”x6” Lintel band");
                subIndex++;

                run.setText(" = " + expressionD.getText().replace("*", "X") + "    ");
                run.setText(" = " + resD.getText() + "cft" + "    ");
            }

            if (eFlag) {
                //Create Another paragraph
                paragraph = document.createParagraph();
                //Set alignment paragraph to RIGHT
                paragraph.setAlignment(ParagraphAlignment.LEFT);
                run = paragraph.createRun();
                run.setText(subIndex + ". Chajja, loft, etc.");
                subIndex++;

                run.setText(" = " + expressionE.getText().replace("*", "X") + "    ");
                run.setText(" = " + resE.getText() + "cft" + "    ");
            }

            if (fFlag) {
                //Create Another paragraph
                paragraph = document.createParagraph();
                //Set alignment paragraph to RIGHT
                paragraph.setAlignment(ParagraphAlignment.LEFT);
                run = paragraph.createRun();
                run.setText(subIndex + ". Roof slab");
                subIndex++;

                run.setText(" = " + expressionF.getText().replace("*", "X") + "    ");
                run.setText(" = " + resF.getText() + "cft" + "    ");
            }

            if (gFlag) {
                //Create Another paragraph
                paragraph = document.createParagraph();
                //Set alignment paragraph to RIGHT
                paragraph.setAlignment(ParagraphAlignment.LEFT);
                run = paragraph.createRun();
                run.setText(subIndex + ". Stair case");
                subIndex++;

                run.setText(" = " + expressionG.getText().replace("*", "X") + "    ");
                run.setText(" = " + resG.getText() + "cft" + "    ");
            }

            if (hFlag) {
                //Create Another paragraph
                paragraph = document.createParagraph();
                //Set alignment paragraph to RIGHT
                paragraph.setAlignment(ParagraphAlignment.LEFT);
                run = paragraph.createRun();
                run.setText(subIndex + ". Front Beam");
                subIndex++;

                run.setText(" = " + expressionH.getText().replace("*", "X") + "    ");
                run.setText(" = " + resH.getText() + "cft" + "    ");
            }

            if (iFlag) {
                //Create Another paragraph
                paragraph = document.createParagraph();
                //Set alignment paragraph to RIGHT
                paragraph.setAlignment(ParagraphAlignment.LEFT);
                run = paragraph.createRun();
                run.setText(subIndex + ". 10”x10” beam on 5” B/W under roof slab");
                subIndex++;

                run.setText(" = " + expressionI.getText().replace("*", "X") + "    ");
                run.setText(" = " + resI.getText() + "cft" + "    ");
            }

            if (jFlag) {
                //Create Another paragraph
                paragraph = document.createParagraph();
                //Set alignment paragraph to RIGHT
                paragraph.setAlignment(ParagraphAlignment.LEFT);
                run = paragraph.createRun();
                run.setText(subIndex + ". Veranda beam ");
                subIndex++;

                run.setText(" = " + expressionJ.getText().replace("*", "X") + "    ");
                run.setText(" = " + resJ.getText() + "cft" + "    ");
            }

            if (kFlag) {
                //Create Another paragraph
                paragraph = document.createParagraph();
                //Set alignment paragraph to RIGHT
                paragraph.setAlignment(ParagraphAlignment.LEFT);
                run = paragraph.createRun();
                run.setText(subIndex + ". R.C.C. railing through roof");
                subIndex++;

                run.setText(" = " + expressionK.getText().replace("*", "X") + "    ");
                run.setText(" = " + resK.getText() + "cft" + "    ");
            }
            paragraph = document.createParagraph();
            run = paragraph.createRun();
            run.setText("-------------------------------------------------------------");

            paragraph = document.createParagraph();
            run = paragraph.createRun();
            run.setText("Total= " + totalSum.getText() + "cft" + "    ");
            run.setText(sumRate.getText() + "%cft" + "    ");
            run.setText(tAmtSum.getText() + "    ");
        }

        if (thckBrkFlag) {
            //4 Create Another paragraph
            paragraph = document.createParagraph();
            //Set alignment paragraph to RIGHT
            paragraph.setAlignment(ParagraphAlignment.LEFT);
            run = paragraph.createRun();
            run.setText(index + ". Supplying & providing 10” brick work in C.M. (1:6)" +
                    " with sone sand and 1st class bricks including all" +
                    " cost of labour & materials, curing etc. completion job");
            index++;

            paragraph = document.createParagraph();
            run = paragraph.createRun();
            run.setText(expression3.getText().replace("*", "X") + "    ");
            run.setText(" =" + result3.getText() + "cft" + "    ");
            run.setText(rate3.getText() + "%cft" + "    ");
            run.setText(tAmt3.getText());
        }

        if (thnBrkFlag) {
            //5 Create Another paragraph
            paragraph = document.createParagraph();
            //Set alignment paragraph to RIGHT
            paragraph.setAlignment(ParagraphAlignment.LEFT);
            run = paragraph.createRun();
            run.setText(index + ". Supplying & providing 5” B/W in C.M. (1:5)" +
                    " with sone-sand and 1st class bricks including all" +
                    " cost of labour & materials, curing etc. completion job");
            index++;
            paragraph = document.createParagraph();
            run = paragraph.createRun();
            run.setText(expression4.getText().replace("*", "X") + "    ");
            run.setText(" =" + result4.getText() + "cft" + "    ");
            run.setText(rate4.getText() + "%cft" + "    ");
            run.setText(tAmt4.getText());
        }

        if (halfFlag) {
            //6 Create Another paragraph
            paragraph = document.createParagraph();
            //Set alignment paragraph to RIGHT
            paragraph.setAlignment(ParagraphAlignment.LEFT);
            run = paragraph.createRun();
            run.setText(index + ". Supplying & providing 1\\2”C.P.in C.M. (1:6) with" +
                    " sone-sand both sides including all cost of" +
                    " labour & materials for both sides");
            index++;

            paragraph = document.createParagraph();
            run = paragraph.createRun();
            run.setText(expression5.getText().replace("*", "X") + "    ");
            run.setText(" =" + result5.getText() + "sft" + "    ");
            run.setText(rate5.getText() + "%cft" + "    ");
            run.setText(tAmt5.getText());
        }

        if (qtrFlag) {
            //7 Create Another paragraph
            paragraph = document.createParagraph();
            //Set alignment paragraph to RIGHT
            paragraph.setAlignment(ParagraphAlignment.LEFT);
            run = paragraph.createRun();
            run.setText(index + ". Supplying & providing 1\\4” C.P in C.M. (1:4) with" +
                    " sone-sand including all cost of labour & materials," +
                    " curing completion job.");
            index++;

            paragraph = document.createParagraph();
            run = paragraph.createRun();
            run.setText(expression6.getText().replace("*", "X") + "    ");
            run.setText(" =" + result6.getText() + "sft" + "    ");
            run.setText(rate6.getText() + "%cft" + "    ");
            run.setText(tAmt6.getText());
        }

        if (steelFlag) {
            //8 Create Another paragraph
            paragraph = document.createParagraph();
            //Set alignment paragraph to RIGHT
            paragraph.setAlignment(ParagraphAlignment.LEFT);
            run = paragraph.createRun();
            run.setText(index + ". Supplying & providing different diameter of steel" +
                    " to the site including all cost of labour & materials," +
                    " transportation cost, taxes etc.");
            index++;

            paragraph = document.createParagraph();
            run = paragraph.createRun();
            run.setText(expression7.getText().replace("*", "X") + "    ");
            run.setText(" =" + result7.getText() + "M.T." + "    ");
            run.setText(rate7.getText() + "P.M.T." + "    ");
            run.setText(tAmt7.getText());
        }

        if (tankFlag) {
            //9 Create Another paragraph
            paragraph = document.createParagraph();
            //Set alignment paragraph to RIGHT
            paragraph.setAlignment(ParagraphAlignment.LEFT);
            run = paragraph.createRun();
            run.setText(index + ". Provision of " + tankSize.getText() + " septic tank" +
                    " with inlet outlet including all cost of" +
                    " labour & materials completion Job.");
            index++;

            paragraph = document.createParagraph();
            run = paragraph.createRun();
            run.setText(" Lump Sum cost = ");
            run.setText(tAmt8.getText());
        }

        if (woodFlag) {
            //10 Create Another paragraph
            paragraph = document.createParagraph();
            //Set alignment paragraph to RIGHT
            paragraph.setAlignment(ParagraphAlignment.LEFT);
            run = paragraph.createRun();
            run.setText(index + ". Supplying & providing with fixing well seasoned" +
                    " sal wood for chaukhat & panel for doors & windows" +
                    " including all cost of labour & materials completion job");
            index++;

            paragraph = document.createParagraph();
            run = paragraph.createRun();
            run.setText(expression9.getText().replace("*", "X") + "    ");
            run.setText(" =" + result9.getText() + "cft" + "    ");
            run.setText(rate9.getText() + "%cft" + "    ");
            run.setText(tAmt9.getText());
        }

        if (ipsFlag) {
            //11 Create Another paragraph
            paragraph = document.createParagraph();
            //Set alignment paragraph to RIGHT
            paragraph.setAlignment(ParagraphAlignment.LEFT);
            run = paragraph.createRun();
            run.setText(index + ". Supplying & Providing 1” I.P.S(1:2:4) with punning" +
                    " On 3” P.C.C.(1:3:6) with sone-sand,chips etc. including" +
                    " all cost of labour & materials completion job");
            index++;

            paragraph = document.createParagraph();
            run = paragraph.createRun();
            run.setText(expression10.getText().replace("*", "X") + "    ");
            run.setText(" =" + result10.getText() + "sft" + "    ");
            run.setText(rate10.getText() + "%sft" + "    ");
            run.setText(tAmt10.getText());

            paragraph = document.createParagraph();
            run = paragraph.createRun();
            run.setText("3” P.C.C.(1:3:6) ");
            run.setText(result10.getText() + "x");
            run.setText(sandRate2A.getText() + "' ");
            run.setText(" =" + result2A.getText() + "cft" + "   ");
            run.setText(rate2A.getText() + "%cft" + "   ");
            run.setText(tAmt2A.getText());
        }

        if (snowFlag) {
            //12 Extra Create Another paragraph
            paragraph = document.createParagraph();
            //Set alignment paragraph to RIGHT
            paragraph.setAlignment(ParagraphAlignment.LEFT);
            run = paragraph.createRun();
            run.setText(index + ". Two coats of snow cement to walls & ceiling" +
                    " including all cost of labour & materials completion Job");
            index++;

            paragraph = document.createParagraph();
            run = paragraph.createRun();
            run.setText(expression11.getText().replace("*", "X") + "    ");
            run.setText(" =" + result11.getText() + "sft" + "    ");
            run.setText(rate11.getText() + "%sft" + "    ");
            run.setText(tAmt11.getText());
        }

        if (paintFlag) {
            //13 Extra Create Another paragraph
            paragraph = document.createParagraph();
            //Set alignment paragraph to RIGHT
            paragraph.setAlignment(ParagraphAlignment.LEFT);
            run = paragraph.createRun();
            run.setText(index + ". Supplying & Providing two coats of enamel paint to chaukhats & pannels" +
                    " including all cost of labour & materials completion Job");
            index++;

            paragraph = document.createParagraph();
            run = paragraph.createRun();
            run.setText(expression12.getText().replace("*", "X") + "    ");
            run.setText(" =" + result12.getText() + "sft" + "    ");
            run.setText(rate12.getText() + "%sft" + "    ");
            run.setText(tAmt12.getText());
        }

        if (graniteFlag) {
            //14 Extra Create Another paragraph
            paragraph = document.createParagraph();
            //Set alignment paragraph to RIGHT
            paragraph.setAlignment(ParagraphAlignment.LEFT);
            run = paragraph.createRun();
            run.setText(index + ". Supplying & Providing " + graniteSize.getText() + " Granite on 3” P.C.C.\n" +
                    " (1:3:6) with sone-sand, chips etc including all\n" +
                    " cost of labors & materials completion Job");
            index++;

            paragraph = document.createParagraph();
            run = paragraph.createRun();
            run.setText(expression13.getText().replace("*", "X") + "    ");
            run.setText(" =" + result13.getText() + "cft" + "    ");
            run.setText(rate13.getText() + "%cft" + "    ");
            run.setText(tAmt13.getText());
        }

        //15 Create Another paragraph
        paragraph = document.createParagraph();
        //Set alignment paragraph to RIGHT
        paragraph.setAlignment(ParagraphAlignment.RIGHT);
        run = paragraph.createRun();
        //run.setBold(true);
        run.setText("TOTAL = ");
        run.setText(String.valueOf(ultimateSum));

        if (xtraFlag) {
            //15 Create Another paragraph
            paragraph = document.createParagraph();
            //Set alignment paragraph to RIGHT
            paragraph.setAlignment(ParagraphAlignment.LEFT);
            run = paragraph.createRun();
            run.setText("EXTRA SERVICE CHARGES");
            if (elecFlag) {
                //Create Another paragraph
                paragraph = document.createParagraph();
                //Set alignment paragraph to RIGHT
                paragraph.setAlignment(ParagraphAlignment.LEFT);
                run = paragraph.createRun();
                run.setText("Add " + elecPerc.getText() + " % for electrical services");
            }
            if (sanitaryFlag) {
                //Create Another paragraph
                paragraph = document.createParagraph();
                //Set alignment paragraph to RIGHT
                paragraph.setAlignment(ParagraphAlignment.LEFT);
                run = paragraph.createRun();
                run.setText("Add " + sanitaryPerc.getText() + " % for sanitary services");
            }
            if (extraFlag) {
                //Create Another paragraph
                paragraph = document.createParagraph();
                //Set alignment paragraph to RIGHT
                paragraph.setAlignment(ParagraphAlignment.LEFT);
                run = paragraph.createRun();
                run.setText("Add " + extraPerc.getText() + " % for extra services");
            }
            paragraph = document.createParagraph();
            run = paragraph.createRun();
            run.setText("-------------------------------------------------------------");

            paragraph = document.createParagraph();
            run = paragraph.createRun();
            run.setText("Total " + totalPerc.getText() + " % charges extra");
        }
        //15 Create Another paragraph
        paragraph = document.createParagraph();
        //Set alignment paragraph to RIGHT
        paragraph.setAlignment(ParagraphAlignment.RIGHT);
        run = paragraph.createRun();
        run.setFontSize(24);
        run.setBold(true);
        run.setText("GRAND TOTAL = ");
        run.setText(String.valueOf(ultimateGrandSum));

        return document;
    }
}
