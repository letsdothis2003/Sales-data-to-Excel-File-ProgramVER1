
import javax.swing.*;
import javax.swing.table.DefaultTableModel;
import java.awt.*;
import java.awt.event.ActionEvent;
import java.awt.event.ActionListener;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.LinkedList;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
//This uses apache to transfer data to excel 
public class SalesExcelProgram extends JFrame {

	/**
	* This program is intended to track of inventory and the sale of products during purchases 
	* The user will input data regarding product name, price and quantity and also customer name.
	* Leave anything blank if you find it unneeded. 
	* 
	* DON'T put letters or symbols for price and quantity.
	* ]t It will trigger an error as those variables are set as integers.
	* I especially mean this for price, I know there's a chance a user will put a dollar sign
	* before that. 
	*/

	private static final long serialVersionUID = 1L;
	private LinkedList<sale> purchaseinfo;
    private JTextField nameinput; //Customer or Cashier Name
    private JTextField productinput; //Product name
    private JTextField priceinput; //Price
    private JTextField quantityleftinput;//Quantity Left
    private JTextField quantitytakeninput;//Quantity taken
    private DefaultTableModel tableModel;
    //In retrospect, I should've added product code and date of sale as variables
    
    
    /**
    * This constructs the user-interface and display along with the table with all the outputs
    */
    public SalesExcelProgram() {
    	purchaseinfo = new LinkedList<>();

        setTitle("Sales Program");
        setSize(400, 400);
        setDefaultCloseOperation(JFrame.EXIT_ON_CLOSE);
        setLayout(new BorderLayout());

        JPanel inputPanel = new JPanel();
        inputPanel.setLayout(new GridLayout(6, 2));

        JLabel nameLabel = new JLabel("Name:");
        nameinput = new JTextField();
        inputPanel.add(nameLabel);
        inputPanel.add(nameinput);

        JLabel productLabel = new JLabel("Product:");
        productinput = new JTextField();
        inputPanel.add(productLabel);
        inputPanel.add(productinput);

        JLabel priceLabel = new JLabel("Price:");
        priceinput = new JTextField();
        inputPanel.add(priceLabel);
        inputPanel.add(priceinput);

        
        JLabel quantityTakenLabel = new JLabel("Quantity Taken:");
        quantitytakeninput = new JTextField();
        inputPanel.add(quantityTakenLabel);
        inputPanel.add(quantitytakeninput);

        JLabel quantityLeftLabel = new JLabel("Quantity Left:");
        quantityleftinput = new JTextField();
        inputPanel.add(quantityLeftLabel);
        inputPanel.add(quantityleftinput);

        JButton addButton = new JButton("Add row");
        addButton.addActionListener(new ActionListener() {
            @Override
            public void actionPerformed(ActionEvent e) {
                addsale();
            }
        });
        inputPanel.add(addButton);

        JButton exportButton = new JButton("Export to Excel");
        exportButton.addActionListener(new ActionListener() {
            @Override
            public void actionPerformed(ActionEvent e) {
                exportToExcel();
            }
        });
        inputPanel.add(exportButton);

        add(inputPanel, BorderLayout.NORTH);

        tableModel = new DefaultTableModel(new String[]{"Name", "Product", "Price", "Quantity Taken", "Quantity Left"}, 0);
        JTable table = new JTable(tableModel);
        JScrollPane scrollPane = new JScrollPane(table);
        add(scrollPane, BorderLayout.CENTER);

        setVisible(true);
    }
    
    /**
    *This adds a row will all the data based on that sale 
    */

    private void addsale() {
        String name = nameinput.getText();
        String product = productinput.getText();
        double price = Double.parseDouble(priceinput.getText());
        int quantityLeft = Integer.parseInt(quantityleftinput.getText());
        int quantityTaken = Integer.parseInt(quantitytakeninput.getText());

        sale newproduct = new sale(name, product, price, quantityTaken, quantityLeft);
        purchaseinfo.add(newproduct);

        Object[] row = {name, product, price, quantityTaken, quantityLeft};
        tableModel.addRow(row);

        clearFields();
    }
    
    /**
    *This will use Apache, a java api and library, to transfer data and create the excel file. 
    */
    
    private void exportToExcel() {
        JFileChooser fileChooser = new JFileChooser();
        fileChooser.setDialogTitle("Save Excel File");
        fileChooser.setFileSelectionMode(JFileChooser.FILES_ONLY);
        int userSelection = fileChooser.showSaveDialog(this);
        if (userSelection == JFileChooser.APPROVE_OPTION) {
            String filePath = fileChooser.getSelectedFile().getAbsolutePath();
            
            // Check if the file path has the .xlsx extension(suffix for all excel files)
            if (!filePath.toLowerCase().endsWith(".xlsx")) {     
                filePath += ".xlsx"; // Append the extension if missing, so user won't have to type it(that would be very annoying)
            }
            
            try (Workbook workbook = new XSSFWorkbook()) {
                Sheet sheet = workbook.createSheet("Todays Stock");

                Row headerRow = sheet.createRow(0);
                headerRow.createCell(0).setCellValue("Name");
                headerRow.createCell(1).setCellValue("Product");
                headerRow.createCell(2).setCellValue("Price");
                headerRow.createCell(3).setCellValue("Quantity Taken");
                headerRow.createCell(4).setCellValue("Quantity Left");

                int rowNum = 1;
                for (sale product : purchaseinfo) {
                    Row row = sheet.createRow(rowNum++);
                    row.createCell(0).setCellValue(product.getName());
                    row.createCell(1).setCellValue(product.getproduct());
                    row.createCell(2).setCellValue(product.getPrice());
                    row.createCell(3).setCellValue(product.getQuantityTaken());
                    row.createCell(4).setCellValue(product.getQuantityLeft());
                }

                try (FileOutputStream outputStream = new FileOutputStream(filePath)) {
                    workbook.write(outputStream);
                    JOptionPane.showMessageDialog(this, "Excel file saved successfully!");
                } catch (IOException e) {
                    e.printStackTrace();
                    JOptionPane.showMessageDialog(this, "Failed to save Excel file, try again or don't!");
                }
            } catch (IOException e) {
                e.printStackTrace();
                JOptionPane.showMessageDialog(this, "Failed to create Excel file, try again or don't!");
            }
        }
    }

    /**
    *This resets after user is done with a row, so it will initialize or remove previous data so user can start on
    *a fresh new row. 
    */
    private void clearFields() {
        nameinput.setText("");
        productinput.setText("");
        priceinput.setText("");
        quantityleftinput.setText("");
        quantitytakeninput.setText("");
        nameinput.requestFocus();
    }

    public static void main(String[] args) {
        SwingUtilities.invokeLater(new Runnable() {
            @Override
            public void run() {
                new SalesExcelProgram();
            }
        });
    }
    

    private static class sale { //Final output
        private String name;
        private String product;
        private double price;
        private int quantityTaken;
        private int quantityLeft;

        public sale(String name, String product, double price, int quantityLeft, int quantityTaken) {
            this.name = name;
            this.product = product;
            this.price = price;
            this.quantityTaken = quantityTaken;
            this.quantityLeft = quantityLeft;
        }

        public String getName() {
            return name;
        }
        public String getproduct() {
            return product;
        }
        public double getPrice() {
            return price;
        }
        public int getQuantityTaken() {
            return quantityTaken;
        }
        public int getQuantityLeft() {
            return quantityLeft;
        }
    }
}
