package principal;

import java.awt.Color;
import java.awt.Component;
import java.awt.Desktop;
import java.awt.EventQueue;

import javax.swing.JFrame;
import javax.swing.JLabel;
import javax.swing.JOptionPane;
import javax.swing.JTextField;
import javax.swing.SwingConstants;
import javax.swing.table.DefaultTableCellRenderer;
import javax.swing.table.DefaultTableModel;
import javax.swing.JScrollPane;
import javax.swing.JButton;
import javax.swing.JTable;


import org.apache.poi.ss.usermodel.BorderStyle;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Header;
import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFPrintSetup;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import javax.print.Doc;
import javax.print.DocFlavor;
import javax.print.DocPrintJob;
import javax.print.PrintException;
import javax.print.PrintService;
import javax.print.PrintServiceLookup;
import javax.print.SimpleDoc;
import javax.swing.ImageIcon;
import java.awt.event.ActionListener;
import java.io.ByteArrayOutputStream;
import java.io.File;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.Map;
import java.util.Set;
import java.util.TreeMap;
import java.awt.event.ActionEvent;

public class Principal extends DefaultTableModel{

	private JFrame frame;
	private JTextField textField;
	private JTextField textField_1;

	/**
	 * Launch the application.
	 */
	public static void main(String[] args) {
		EventQueue.invokeLater(new Runnable() {
			public void run() {
				try {
					Principal window = new Principal();
					window.frame.setVisible(true);
					window.frame.setLocationRelativeTo(null);
				} catch (Exception e) {
					e.printStackTrace();
				}
			}
		});
	}

	/**
	 * Create the application.
	 */
	public Principal() {
		initialize();
	}

	/**
	 * Initialize the contents of the frame.
	 */
	private void initialize(){
		

		String [] colunas = {"Segunda", "Treça", "Quarta", "Quinta", "Sexta", "Sabado"};
		
		Object [][] dados = {
			    {"", "", "", "", "", "", ""},
			    {"", "", "", "", "", "", ""},
			    {"", "", "", "", "", "", ""},
			    {"", "", "", "", "", "", ""},
			    {"", "", "", "", "", "", ""},
			    {"", "", "", "", "", "", ""},
			    {"", "", "", "", "", "", ""},
			    {"###############", "###############", "###############", "###############", "###############", "###############", "###############"},
			    {"", "", "", "", "", "", ""},
			    {"", "", "", "", "", "", ""},
			    {"", "", "", "", "", "", ""},
			    {"", "", "", "", "", "", ""},
			    {"", "", "", "", "", "", ""},
			    {"", "", "", "", "", "", ""},
			    {"", "", "", "", "", "", ""},
			    {"###############", "###############", "###############", "###############", "###############", "###############", "###############"},
			    {"", "", "", "", "", "", ""},
			    {"", "", "", "", "", "", ""},
			    {"", "", "", "", "", "", ""},
			    {"", "", "", "", "", "", ""},
			    {"", "", "", "", "", "", ""},
			    {"", "", "", "", "", "", ""},
			    {"", "", "", "", "", "", ""},
			    {"###############", "###############", "###############", "###############", "###############", "###############", "###############"},			    
			    {"", "", "", "", "", "", ""},
			    {"", "", "", "", "", "", ""},
			    {"", "", "", "", "", "", ""},
			    {"", "", "", "", "", "", ""},
			    {"", "", "", "", "", "", ""},
			    {"", "", "", "", "", "", ""},
			    {"", "", "", "", "", "", ""},
			};
		
		frame = new JFrame();
		frame.setUndecorated(true);
		frame.setBounds(100, 100, 715, 651);
		frame.setDefaultCloseOperation(JFrame.EXIT_ON_CLOSE);
		frame.getContentPane().setLayout(null);
		
		JLabel label = new JLabel("M\u00EAs");
		label.setBounds(337, 11, 30, 14);
		frame.getContentPane().add(label);
		
		textField = new JTextField();
		textField.setHorizontalAlignment(SwingConstants.CENTER);
		textField.setColumns(10);
		textField.setBounds(283, 24, 130, 20);
		frame.getContentPane().add(textField);
		
		JLabel label_1 = new JLabel("");
		label_1.setIcon(new ImageIcon("C:\\Users\\u6071072\\eclipse-workspace\\Nutrisul\\cozinha.industrial.nutrisul.6137.jpg.png"));
		label_1.setBounds(10, 11, 70, 33);
		frame.getContentPane().add(label_1);
		
		JButton button = new JButton("Sair");
		button.addActionListener(new ActionListener() {
			public void actionPerformed(ActionEvent arg0) {
				System.exit(0);
			}
		});
		button.setBounds(109, 598, 59, 23);
		frame.getContentPane().add(button);
		
		
		JTable tabela = new JTable(dados,colunas);
		tabela.setDefaultRenderer(Object.class, new CellRenderer());
		JScrollPane scroll = new JScrollPane();
		scroll.setBounds(10, 55, 679, 523);
		scroll.setViewportView(tabela);
		frame.getContentPane().add(scroll);
		
		JButton button_1 = new JButton("Imprimir");
		button_1.addActionListener(new ActionListener() {
			public void actionPerformed(ActionEvent e){

				
				String Mes = textField.getText();
				XSSFWorkbook workbook = new XSSFWorkbook();

			      FileOutputStream out = null;
				try {
					out = new FileOutputStream(new File("Cardapio.xlsx"));
				} catch (FileNotFoundException e1) {
					// TODO Auto-generated catch block
					JOptionPane.showMessageDialog(null,"Arquivo Aberto. Feche para continuar a impressão!", null, JOptionPane.ERROR_MESSAGE);
					e1.printStackTrace();
				}
			      XSSFSheet spreadsheet = workbook.createSheet("Cardapio");
			      Header header = spreadsheet.getHeader();		
			      header.setCenter("Cardápio Mês " +Mes);
			      workbook.setPrintArea(0, 0, 5, 0, 6);
			      spreadsheet.setColumnWidth((short) (0), (short) (10 * 612.55));
			      spreadsheet.setColumnWidth((short) (1), (short) (10 * 612.55));
			      spreadsheet.setColumnWidth((short) (2), (short) (10 * 612.55));
			      spreadsheet.setColumnWidth((short) (3), (short) (10 * 612.55));
			      spreadsheet.setColumnWidth((short) (4), (short) (10 * 612.55));
			      spreadsheet.setColumnWidth((short) (5), (short) (10 * 612.55));
			      XSSFRow row;
			      
			      XSSFCellStyle estilo = workbook.createCellStyle();
			      estilo.setAlignment(HorizontalAlignment.CENTER);
	              estilo.setBorderLeft(BorderStyle.THIN);
	              estilo.setBorderBottom(BorderStyle.THIN);
	              estilo.setBorderRight(BorderStyle.THIN);
	              estilo.setBorderTop(BorderStyle.THIN);
	              
			      

			      Map < String, Object[] > empinfo = 
			      new TreeMap < String, Object[] >();
			      empinfo.put( "1", new Object[] { "Segunda", "Terça", "Quarta", "Quinta", "Sexta", "Sabado" });
			      empinfo.put( "11", new Object[] { ""+tabela.getValueAt(0, 0), ""+tabela.getValueAt(0, 1), ""+tabela.getValueAt(0, 2), ""+tabela.getValueAt(0, 3), ""+tabela.getValueAt(0, 4), ""+tabela.getValueAt(0, 5) });
			      empinfo.put( "111", new Object[] { ""+tabela.getValueAt(1, 0), ""+tabela.getValueAt(1, 1), ""+tabela.getValueAt(1, 2), ""+tabela.getValueAt(1, 3), ""+tabela.getValueAt(1, 4), ""+tabela.getValueAt(1, 5) });
			      empinfo.put( "1111", new Object[] { ""+tabela.getValueAt(2, 0), ""+tabela.getValueAt(2, 1), ""+tabela.getValueAt(2, 2), ""+tabela.getValueAt(2, 3), ""+tabela.getValueAt(2, 4), ""+tabela.getValueAt(2, 5) });
			      empinfo.put( "11111", new Object[] { ""+tabela.getValueAt(3, 0), ""+tabela.getValueAt(3, 1), ""+tabela.getValueAt(3, 2), ""+tabela.getValueAt(3, 3), ""+tabela.getValueAt(3, 4), ""+tabela.getValueAt(3, 5) });
			      empinfo.put( "111111", new Object[] { ""+tabela.getValueAt(4, 0), ""+tabela.getValueAt(4, 1), ""+tabela.getValueAt(4, 2), ""+tabela.getValueAt(4, 3), ""+tabela.getValueAt(4, 4), ""+tabela.getValueAt(4, 5) });
			      empinfo.put( "1111111", new Object[] { ""+tabela.getValueAt(5, 0), ""+tabela.getValueAt(5, 1), ""+tabela.getValueAt(5, 2), ""+tabela.getValueAt(5, 3), ""+tabela.getValueAt(5, 4), ""+tabela.getValueAt(5, 5) });
			      empinfo.put( "11111111", new Object[] { ""+tabela.getValueAt(6, 0), ""+tabela.getValueAt(6, 1), ""+tabela.getValueAt(6, 2), ""+tabela.getValueAt(6, 3), ""+tabela.getValueAt(6, 4), ""+tabela.getValueAt(6, 5) });
			      empinfo.put( "111111111", new Object[] { "Segunda", "Terça", "Quarta", "Quinta", "Sexta", "Sabado" });
			      empinfo.put( "1111111111", new Object[] { ""+tabela.getValueAt(8, 0), ""+tabela.getValueAt(8, 1), ""+tabela.getValueAt(8, 2), ""+tabela.getValueAt(8, 3), ""+tabela.getValueAt(8, 4), ""+tabela.getValueAt(8, 5) });
			      empinfo.put( "11111111111", new Object[] { ""+tabela.getValueAt(9, 0), ""+tabela.getValueAt(9, 1), ""+tabela.getValueAt(9, 2), ""+tabela.getValueAt(9, 3), ""+tabela.getValueAt(9, 4), ""+tabela.getValueAt(9, 5) });
			      empinfo.put( "111111111111", new Object[] { ""+tabela.getValueAt(10, 0), ""+tabela.getValueAt(10, 1), ""+tabela.getValueAt(10, 2), ""+tabela.getValueAt(10, 3), ""+tabela.getValueAt(10, 4), ""+tabela.getValueAt(10, 5) });
			      empinfo.put( "1111111111111", new Object[] { ""+tabela.getValueAt(11, 0), ""+tabela.getValueAt(11, 1), ""+tabela.getValueAt(11, 2), ""+tabela.getValueAt(11, 3), ""+tabela.getValueAt(11, 4), ""+tabela.getValueAt(11, 5) });
			      empinfo.put( "11111111111111", new Object[] { ""+tabela.getValueAt(12, 0), ""+tabela.getValueAt(12, 1), ""+tabela.getValueAt(12, 2), ""+tabela.getValueAt(12, 3), ""+tabela.getValueAt(12, 4), ""+tabela.getValueAt(12, 5) });
			      empinfo.put( "111111111111111", new Object[] { ""+tabela.getValueAt(13, 0), ""+tabela.getValueAt(13, 1), ""+tabela.getValueAt(13, 2), ""+tabela.getValueAt(13, 3), ""+tabela.getValueAt(13, 4), ""+tabela.getValueAt(13, 5) });
			      empinfo.put( "1111111111111111", new Object[] { ""+tabela.getValueAt(14, 0), ""+tabela.getValueAt(14, 1), ""+tabela.getValueAt(14, 2), ""+tabela.getValueAt(14, 3), ""+tabela.getValueAt(14, 4), ""+tabela.getValueAt(14, 5) });
			      empinfo.put( "11111111111111111", new Object[] { "Segunda", "Terça", "Quarta", "Quinta", "Sexta", "Sabado" });
			      empinfo.put( "111111111111111111", new Object[] { ""+tabela.getValueAt(16, 0), ""+tabela.getValueAt(16, 1), ""+tabela.getValueAt(16, 2), ""+tabela.getValueAt(16, 3), ""+tabela.getValueAt(16, 4), ""+tabela.getValueAt(16, 5) });
			      empinfo.put( "1111111111111111111", new Object[] { ""+tabela.getValueAt(17, 0), ""+tabela.getValueAt(17, 1), ""+tabela.getValueAt(17, 2), ""+tabela.getValueAt(17, 3), ""+tabela.getValueAt(17, 4), ""+tabela.getValueAt(17, 5) });
			      empinfo.put( "11111111111111111111", new Object[] { ""+tabela.getValueAt(18, 0), ""+tabela.getValueAt(18, 1), ""+tabela.getValueAt(18, 2), ""+tabela.getValueAt(18, 3), ""+tabela.getValueAt(18, 4), ""+tabela.getValueAt(18, 5) });
			      empinfo.put( "111111111111111111111", new Object[] { ""+tabela.getValueAt(19, 0), ""+tabela.getValueAt(19, 1), ""+tabela.getValueAt(19, 2), ""+tabela.getValueAt(19, 3), ""+tabela.getValueAt(19, 4), ""+tabela.getValueAt(19, 5) });
			      empinfo.put( "1111111111111111111111", new Object[] { ""+tabela.getValueAt(20, 0), ""+tabela.getValueAt(20, 1), ""+tabela.getValueAt(20, 2), ""+tabela.getValueAt(20, 3), ""+tabela.getValueAt(20, 4), ""+tabela.getValueAt(20, 5) });
			      empinfo.put( "11111111111111111111111", new Object[] { ""+tabela.getValueAt(21, 0), ""+tabela.getValueAt(21, 1), ""+tabela.getValueAt(21, 2), ""+tabela.getValueAt(21, 3), ""+tabela.getValueAt(21, 4), ""+tabela.getValueAt(21, 5) });
			      empinfo.put( "111111111111111111111111", new Object[] { ""+tabela.getValueAt(22, 0), ""+tabela.getValueAt(22, 1), ""+tabela.getValueAt(22, 2), ""+tabela.getValueAt(22, 3), ""+tabela.getValueAt(22, 4), ""+tabela.getValueAt(22, 5) });
			      empinfo.put( "1111111111111111111111111", new Object[] { "Segunda", "Terça", "Quarta", "Quinta", "Sexta", "Sabado" });
			      empinfo.put( "11111111111111111111111111", new Object[] { ""+tabela.getValueAt(24, 0), ""+tabela.getValueAt(24, 1), ""+tabela.getValueAt(24, 2), ""+tabela.getValueAt(24, 3), ""+tabela.getValueAt(24, 4), ""+tabela.getValueAt(24, 5) });
			      empinfo.put( "111111111111111111111111111", new Object[] { ""+tabela.getValueAt(25, 0), ""+tabela.getValueAt(25, 1), ""+tabela.getValueAt(25, 2), ""+tabela.getValueAt(25, 3), ""+tabela.getValueAt(25, 4), ""+tabela.getValueAt(25, 5) });
			      empinfo.put( "1111111111111111111111111111", new Object[] { ""+tabela.getValueAt(26, 0), ""+tabela.getValueAt(26, 1), ""+tabela.getValueAt(26, 2), ""+tabela.getValueAt(26, 3), ""+tabela.getValueAt(26, 4), ""+tabela.getValueAt(26, 5) });
			      empinfo.put( "11111111111111111111111111111", new Object[] { ""+tabela.getValueAt(27, 0), ""+tabela.getValueAt(27, 1), ""+tabela.getValueAt(27, 2), ""+tabela.getValueAt(27, 3), ""+tabela.getValueAt(27, 4), ""+tabela.getValueAt(27, 5) });
			      empinfo.put( "111111111111111111111111111111", new Object[] { ""+tabela.getValueAt(28, 0), ""+tabela.getValueAt(28, 1), ""+tabela.getValueAt(28, 2), ""+tabela.getValueAt(28, 3), ""+tabela.getValueAt(28, 4), ""+tabela.getValueAt(28, 5) });
			      empinfo.put( "1111111111111111111111111111111", new Object[] { ""+tabela.getValueAt(29, 0), ""+tabela.getValueAt(29, 1), ""+tabela.getValueAt(29, 2), ""+tabela.getValueAt(29, 3), ""+tabela.getValueAt(29, 4), ""+tabela.getValueAt(29, 5) });
			      empinfo.put( "11111111111111111111111111111111", new Object[] { ""+tabela.getValueAt(30, 0), ""+tabela.getValueAt(30, 1), ""+tabela.getValueAt(30, 2), ""+tabela.getValueAt(30, 3), ""+tabela.getValueAt(30, 4), ""+tabela.getValueAt(30, 5) });
			      empinfo.put( "111111111111111111111111111111111", new Object[] { ""+textField_1.getText() ,"","","","",""});
			      
			      Set < String > keyid = empinfo.keySet();
			      int rowid = 0;
			      for (String key : keyid) {
			         row = spreadsheet.createRow(rowid++);
			         Object [] objectArr = empinfo.get(key);
			         int cellid = 0;

			         for (Object obj : objectArr) {
			            Cell cell = row.createCell(cellid++);
			            cell.setCellValue((String)obj);
			            	if(obj == "Segunda" || obj == "Terça" || obj == "Quarta"|| obj == "Quinta"|| obj == "Sexta"|| obj == "Sabado") {
			            		
			            	}else {
			            		cell.setCellStyle(estilo);
			            	}
			            			            
			         }
			      }
			      
			      spreadsheet.addMergedRegion(new CellRangeAddress(32,32,0,5));
			      workbook.setPrintArea(0, 0, 6, 0, 32);
			      spreadsheet.getPrintSetup().setPaperSize(XSSFPrintSetup.A4_PAPERSIZE);
			      spreadsheet.setDisplayGridlines(true);
			      spreadsheet.getPrintSetup().setLandscape(true);
			      spreadsheet.setMargin(Sheet.LeftMargin, 0.2);
			      spreadsheet.setMargin(Sheet.RightMargin, 0.2);
			           
			      
			      //Imprime o arquivo Excel
			      PrintService printService = PrintServiceLookup.lookupDefaultPrintService();
			      ByteArrayOutputStream bos = new ByteArrayOutputStream();
			      
			       				
			       byte[] by = bos.toByteArray();
			       DocFlavor flavor = DocFlavor.BYTE_ARRAY.AUTOSENSE;
			       Doc doc = new SimpleDoc(by, flavor, null);  
			       DocPrintJob job = printService.createPrintJob();
			      try {
						job.print(doc, null);
					} catch (PrintException e2) {
						JOptionPane.showMessageDialog(null,"Verifique a Impressora!", null, JOptionPane.ERROR_MESSAGE);
					}
			      
			      
			      //Gera o Arquivo Excel
			      	try {
			      		java.awt.Desktop.getDesktop().open(new File("Cardapio.xlsx"));
			      	} catch (IOException e1) {
			      		e1.printStackTrace();
			      	} 
			      	
			      	
			      	try {
			    	  workbook.write(out);	
			    	  out.close();
					} catch (IOException e2) {
					e2.printStackTrace();
					}
			      	
			      	
			      System.out.println("Salvo com Sucesso");
			      
				
			}
		});
		button_1.setBounds(10, 598, 89, 23);
		
		frame.getContentPane().add(button_1);
		
		JLabel lblObs = new JLabel("OBS:");
		lblObs.setBounds(379, 602, 46, 14);
		frame.getContentPane().add(lblObs);
		
		textField_1 = new JTextField();
		textField_1.setBounds(411, 599, 278, 20);
		frame.getContentPane().add(textField_1);
		textField_1.setColumns(10);
		
		
	}
	public class CellRenderer extends DefaultTableCellRenderer {
		public CellRenderer() {
			super();
		}
		public Component getTableCellRendererComponent(JTable table, Object value,
				boolean isSelected, boolean hasFocus, int row, int column) {
			this.setHorizontalAlignment(CENTER);
			return super.getTableCellRendererComponent(table, value, isSelected,
					hasFocus, row, column);
		}
	}
}

