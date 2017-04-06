/**
 * Created by vinee on 31-Mar-17.
 */

import javafx.application.Application;
import javafx.event.ActionEvent;
import javafx.event.EventHandler;
import javafx.geometry.HPos;
import javafx.geometry.Pos;
import javafx.geometry.VPos;
import javafx.scene.Scene;
import javafx.scene.control.*;
import javafx.scene.layout.GridPane;
import javafx.scene.text.Text;
import javafx.stage.FileChooser;
import javafx.stage.Stage;
import net.objecthunter.exp4j.Expression;
import net.objecthunter.exp4j.ExpressionBuilder;
import org.apache.poi.xwpf.usermodel.ParagraphAlignment;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFRun;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;

public class FeetCalculator extends Application {
    private static final String defaultFileName = "";
    CheckBox cb1 = new CheckBox("E/W in excavation of soil");
    CheckBox cb2 = new CheckBox("world");
    double res;
    private Stage savedStage;
    private boolean flag2 = false;
    private TextField custName = new TextField();
    private TextField Feet = new TextField();
    private TextField inches = new TextField();
    private TextField totalFeet = new TextField();
    private TextField expression = new TextField();
    private Text result = new Text();
    private TextField rate = new TextField();

    private Text tAmt = new Text();
    private Button btConvert = new Button("Convert");
    private Button btCalculateQty = new Button("Calculate Qty.");
    private Button btAmount = new Button("Calculate Amt.");
    private Button generate = new Button("Generate");

    private Text actionTarget = new Text("");

    @Override // Override the start method in the Application class
    public void start(Stage primaryStage) {
        ScrollPane scrollPane = new ScrollPane();
        GridPane gridPane = new GridPane();

//        RowConstraints row = new RowConstraints();
//        row.setPercentHeight(100);
//        row.setFillHeight(false);
//        row.setValignment(VPos.CENTER);
//
//        ColumnConstraints col = new ColumnConstraints();
//        col.setPercentWidth(100);
//        col.setFillWidth(false);
//        col.setHalignment(HPos.CENTER);

        // Create UI
        gridPane.setHgap(5);
        gridPane.setVgap(5);
        gridPane.add(new Label("Enter feet"), 0, 0);
        gridPane.add(Feet, 1, 0);
        gridPane.add(new Label("Enter inches"), 0, 1);
        gridPane.add(inches, 1, 1);
        gridPane.add(new Label("Total feet:"), 0, 2);
        gridPane.add(totalFeet, 1, 2);
        gridPane.add(btConvert, 1, 3);

        gridPane.add(new Label("Enter Expression:"), 0, 5);
        gridPane.add(expression, 1, 5);
        gridPane.add(new Label("Result:"), 2, 5);
        gridPane.add(result, 3, 5);
        gridPane.add(new Label("Enter Rate %:"), 4, 5);
        gridPane.add(rate, 5, 5);
        gridPane.add(new Label("Amount:"), 6, 5);
        gridPane.add(tAmt, 7, 5);
        gridPane.add(btCalculateQty, 1, 6);
        gridPane.add(btAmount, 5, 6);

        gridPane.add(cb1, 0, 4);
        gridPane.add(cb2, 1, 7);

        gridPane.add(generate, 0, 8);
        gridPane.add(actionTarget, 0, 9);
        gridPane.add(custName, 0, 10);

        final Separator sepHor = new Separator();
        sepHor.setValignment(VPos.CENTER);
        GridPane.setConstraints(sepHor, 1, 4);
        GridPane.setColumnSpan(sepHor, 10);
        gridPane.getChildren().add(sepHor);

        // Set properties for UI
        gridPane.setAlignment(Pos.CENTER);
        Feet.setAlignment(Pos.BOTTOM_RIGHT);
        inches.setAlignment(Pos.BOTTOM_RIGHT);
        totalFeet.setAlignment(Pos.BOTTOM_RIGHT);
        btConvert.setAlignment(Pos.BOTTOM_RIGHT);
        expression.setAlignment(Pos.BOTTOM_RIGHT);
        //result.setAlignment(Pos.BOTTOM_RIGHT);

        totalFeet.setEditable(false);
        //result.setEditable(false);
        GridPane.setHalignment(btConvert, HPos.RIGHT);

        scrollPane.setHbarPolicy(ScrollPane.ScrollBarPolicy.ALWAYS);
        scrollPane.setVbarPolicy(ScrollPane.ScrollBarPolicy.ALWAYS);
        scrollPane.setContent(gridPane);

        // Process events
        btConvert.setOnAction(event -> Convert());
        btCalculateQty.setOnAction(event -> calculateQty());
        btAmount.setOnAction(event -> calculateAmount());
        generate.setOnAction(event -> {
            try {
                Generate();
            } catch (IOException e) {
                e.printStackTrace();
            }
        });

        cb1.setOnAction(new EventHandler<ActionEvent>() {
            @Override
            public void handle(ActionEvent event) {
                System.out.println("CB1 ticked !");
            }
        });
        cb2.setOnAction(new EventHandler<ActionEvent>() {
            @Override
            public void handle(ActionEvent event) {
                System.out.println("CB2 ticked!");
                flag2 = true;
            }
        });

        // Create a scene and place it in the stage
        Scene scene = new Scene(scrollPane, 1024, 768);
        primaryStage.setTitle("Home Estimate"); // Set title
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

    private void calculateAmount() {
        double amountVal = (res * Double.valueOf(rate.getText()) / 100);
        tAmt.setText(String.valueOf(amountVal));
    }

    private void calculateQty() {
        // Get values from text fields

        Expression e = new ExpressionBuilder(expression.getText()).build();
        res = e.evaluate();

        result.setText(String.valueOf(res));
    }

    private void Generate() throws IOException {
        actionTarget.setText("Document generated");
        showSaveFileChooser();
    }

    private void showSaveFileChooser() throws IOException {
        FileChooser fileChooser = new FileChooser();
        fileChooser.setTitle("Save file");
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

    private XWPFDocument myFileWriter() throws IOException {
        //Blank Document
        XWPFDocument document = new XWPFDocument();

        //create paragraph
        XWPFParagraph paragraph = document.createParagraph();

        String name = custName.getText();

        //Set alignment paragraph to RIGHT
        paragraph.setAlignment(ParagraphAlignment.CENTER);
        XWPFRun run = paragraph.createRun();
        run.setBold(true);
        run.setFontSize(24);
        run.setText("Estimated cost of residential building extention/ renovation G.F.&F.F) of" +
                custName.getText() + " W/O Mr. Kamlesh baitha Khata-76 Plot No-386, Muhalla-Agarwa" +
                " Tauji No-51 MOTIHARI (EastChamparan)");

        if (flag2) {
            //Create Another paragraph
            paragraph = document.createParagraph();
            //Set alignment paragraph to RIGHT
            paragraph.setAlignment(ParagraphAlignment.CENTER);
            run = paragraph.createRun();
            run.setText(" BUILT UP G.F AREA-552 SFT");
        }
        return document;
    }
}