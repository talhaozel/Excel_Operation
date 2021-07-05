
package thyproject;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.Vector;
import java.util.logging.Level;
import java.util.logging.Logger;
import javax.swing.ImageIcon;
import javax.swing.JFileChooser;
import javax.swing.JOptionPane;
import javax.swing.RowFilter;
import javax.swing.table.DefaultTableModel;
import javax.swing.table.TableRowSorter;
import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import static org.apache.poi.ss.usermodel.CellType.BOOLEAN;
import static org.apache.poi.ss.usermodel.CellType.NUMERIC;
import static org.apache.poi.ss.usermodel.CellType.STRING;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
/**
 *
 * @author Talha
 */

public class AnaSayfa extends javax.swing.JFrame {   
   
    DefaultTableModel model ; 

    ArrayList<Object> arr = new ArrayList<Object>();
 
    public AnaSayfa() throws FileNotFoundException, IOException {
        initComponents();
  
        jLabel1.setIcon(new ImageIcon("C:\\Users\\Talha\\Desktop\\icon.png"));
    
    }

    @SuppressWarnings("unchecked")
    // <editor-fold defaultstate="collapsed" desc="Generated Code">//GEN-BEGIN:initComponents
    private void initComponents() {

        jScrollPane1 = new javax.swing.JScrollPane();
        jTable1 = new javax.swing.JTable();
        btn_update = new javax.swing.JButton();
        btn_excel = new javax.swing.JButton();
        txtSearch = new javax.swing.JTextField();
        jLabel1 = new javax.swing.JLabel();
        updText = new javax.swing.JTextField();
        dosya_yol = new javax.swing.JTextField();
        jLabel2 = new javax.swing.JLabel();

        setDefaultCloseOperation(javax.swing.WindowConstants.EXIT_ON_CLOSE);
        setTitle("TURKISH AIRLINES");
        setBackground(new java.awt.Color(255, 0, 0));
        setForeground(java.awt.Color.white);
        setMinimumSize(new java.awt.Dimension(1000, 800));

        jTable1.setBackground(new java.awt.Color(255, 255, 255));
        jTable1.setForeground(new java.awt.Color(0, 0, 0));
        jTable1.setMinimumSize(new java.awt.Dimension(1300, 600));
        jTable1.addMouseListener(new java.awt.event.MouseAdapter() {
            public void mouseClicked(java.awt.event.MouseEvent evt) {
                jTable1MouseClicked(evt);
            }
        });
        jScrollPane1.setViewportView(jTable1);

        btn_update.setFont(new java.awt.Font("Segoe UI", 1, 12)); // NOI18N
        btn_update.setText("UPDATE");
        btn_update.setCursor(new java.awt.Cursor(java.awt.Cursor.DEFAULT_CURSOR));
        btn_update.addActionListener(new java.awt.event.ActionListener() {
            public void actionPerformed(java.awt.event.ActionEvent evt) {
                btn_updateActionPerformed(evt);
            }
        });

        btn_excel.setFont(new java.awt.Font("Segoe UI", 1, 12)); // NOI18N
        btn_excel.setText("Choose excel file");
        btn_excel.addActionListener(new java.awt.event.ActionListener() {
            public void actionPerformed(java.awt.event.ActionEvent evt) {
                btn_excelActionPerformed(evt);
            }
        });

        txtSearch.setText("Searching word");
        txtSearch.addKeyListener(new java.awt.event.KeyAdapter() {
            public void keyTyped(java.awt.event.KeyEvent evt) {
                txtSearchKeyTyped(evt);
            }
        });

        jLabel1.setHorizontalAlignment(javax.swing.SwingConstants.CENTER);

        updText.setText("Update data");

        dosya_yol.setText("File path");
        dosya_yol.setToolTipText("");
        dosya_yol.addActionListener(new java.awt.event.ActionListener() {
            public void actionPerformed(java.awt.event.ActionEvent evt) {
                dosya_yolActionPerformed(evt);
            }
        });

        jLabel2.setIcon(new javax.swing.ImageIcon(getClass().getResource("/thyproject/search.png"))); // NOI18N

        javax.swing.GroupLayout layout = new javax.swing.GroupLayout(getContentPane());
        getContentPane().setLayout(layout);
        layout.setHorizontalGroup(
            layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
            .addGroup(layout.createSequentialGroup()
                .addGroup(layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
                    .addGroup(layout.createSequentialGroup()
                        .addContainerGap()
                        .addComponent(jScrollPane1))
                    .addGroup(layout.createSequentialGroup()
                        .addGap(50, 50, 50)
                        .addGroup(layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
                            .addComponent(updText, javax.swing.GroupLayout.PREFERRED_SIZE, 274, javax.swing.GroupLayout.PREFERRED_SIZE)
                            .addComponent(dosya_yol, javax.swing.GroupLayout.PREFERRED_SIZE, 274, javax.swing.GroupLayout.PREFERRED_SIZE)
                            .addComponent(txtSearch, javax.swing.GroupLayout.PREFERRED_SIZE, 274, javax.swing.GroupLayout.PREFERRED_SIZE))
                        .addPreferredGap(javax.swing.LayoutStyle.ComponentPlacement.RELATED)
                        .addGroup(layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
                            .addGroup(layout.createSequentialGroup()
                                .addGroup(layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
                                    .addComponent(btn_update, javax.swing.GroupLayout.PREFERRED_SIZE, 122, javax.swing.GroupLayout.PREFERRED_SIZE)
                                    .addComponent(jLabel2, javax.swing.GroupLayout.PREFERRED_SIZE, 114, javax.swing.GroupLayout.PREFERRED_SIZE))
                                .addPreferredGap(javax.swing.LayoutStyle.ComponentPlacement.RELATED, javax.swing.GroupLayout.DEFAULT_SIZE, Short.MAX_VALUE)
                                .addComponent(jLabel1, javax.swing.GroupLayout.PREFERRED_SIZE, 434, javax.swing.GroupLayout.PREFERRED_SIZE)
                                .addGap(15, 15, 15))
                            .addGroup(layout.createSequentialGroup()
                                .addComponent(btn_excel)
                                .addGap(0, 0, Short.MAX_VALUE)))))
                .addContainerGap())
        );
        layout.setVerticalGroup(
            layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
            .addGroup(layout.createSequentialGroup()
                .addGap(22, 22, 22)
                .addGroup(layout.createParallelGroup(javax.swing.GroupLayout.Alignment.BASELINE)
                    .addComponent(dosya_yol, javax.swing.GroupLayout.PREFERRED_SIZE, javax.swing.GroupLayout.DEFAULT_SIZE, javax.swing.GroupLayout.PREFERRED_SIZE)
                    .addComponent(btn_excel))
                .addGroup(layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
                    .addGroup(layout.createSequentialGroup()
                        .addGap(24, 24, 24)
                        .addComponent(jLabel1, javax.swing.GroupLayout.PREFERRED_SIZE, 86, javax.swing.GroupLayout.PREFERRED_SIZE))
                    .addGroup(layout.createSequentialGroup()
                        .addGap(37, 37, 37)
                        .addGroup(layout.createParallelGroup(javax.swing.GroupLayout.Alignment.BASELINE)
                            .addComponent(updText, javax.swing.GroupLayout.PREFERRED_SIZE, javax.swing.GroupLayout.DEFAULT_SIZE, javax.swing.GroupLayout.PREFERRED_SIZE)
                            .addComponent(btn_update, javax.swing.GroupLayout.PREFERRED_SIZE, 22, javax.swing.GroupLayout.PREFERRED_SIZE))
                        .addGap(32, 32, 32)
                        .addGroup(layout.createParallelGroup(javax.swing.GroupLayout.Alignment.TRAILING)
                            .addComponent(txtSearch, javax.swing.GroupLayout.PREFERRED_SIZE, javax.swing.GroupLayout.DEFAULT_SIZE, javax.swing.GroupLayout.PREFERRED_SIZE)
                            .addComponent(jLabel2))))
                .addPreferredGap(javax.swing.LayoutStyle.ComponentPlacement.RELATED, 74, Short.MAX_VALUE)
                .addComponent(jScrollPane1, javax.swing.GroupLayout.PREFERRED_SIZE, 390, javax.swing.GroupLayout.PREFERRED_SIZE)
                .addContainerGap(629, Short.MAX_VALUE))
        );

        pack();
    }// </editor-fold>//GEN-END:initComponents

    private void btn_updateActionPerformed(java.awt.event.ActionEvent evt) {//GEN-FIRST:event_btn_updateActionPerformed
  
         String excelFilePath = dosya_yol.getText();
         FileInputStream inputStream;
                try {
                    inputStream = new FileInputStream(new File(excelFilePath));
                    
                    Workbook workbook =WorkbookFactory.create(inputStream);
        
                    Sheet sheet = workbook.getSheetAt(0);
                    Cell cell2Update = sheet.getRow(jTable1.getSelectedRow()+1).getCell(jTable1.getSelectedColumn());
                    cell2Update.setCellValue(updText.getText());

                    FileOutputStream outputStream = new FileOutputStream(excelFilePath);
                      
                        workbook.write(outputStream);
                        workbook.close();
                        outputStream.close();

                    int temp = ((jTable1.getSelectedRow()+1) *( jTable1.getColumnCount())+ (jTable1.getSelectedColumn()));
                   
                    arr.set(temp,  updText.getText());

                    DefaultTableModel updModel = new DefaultTableModel();

           try{
                     
                    FileInputStream inputstream=new FileInputStream(excelFilePath);
                                 
		
                    XSSFWorkbook workbook2 =new XSSFWorkbook(inputstream);
                    XSSFSheet sheet2 =workbook2.getSheetAt(0);	

                    int rows=sheet2.getLastRowNum();
                    int cols=sheet2.getRow(1).getLastCellNum();
		 
                
              
                    for(int r=0;r<=rows;r++)
                    {
			XSSFRow row=sheet2.getRow(r); 
			                      
			for(int c=0;c<cols;c++)
			{
                            
                            XSSFCell cell=row.getCell(c,org.apache.poi.ss.usermodel.Row.MissingCellPolicy.CREATE_NULL_AS_BLANK);
                          
                                  
                            if(r==0){
                                 
                             updModel.addColumn(row.getCell(c));
                            }
                                                              
                              CellType type = cell.getCellType();
                                    
                        }
                    } 
                                      
                        for(int i = 0; i< arr.size() ; i=i+cols)
                        {

                            if(i==0){                     
                                continue;
                            }   
                             else
                            {

                             Vector vector = new Vector();

                             for(int k=0; k< cols; k++){

                               vector.add(arr.get(i+k));
                            }

                        updModel.addRow(vector);

                            }
                        }
                }  
                catch (Exception x)
                {
                    x.printStackTrace();
                }                   
                    jTable1.removeAll();
                    jTable1.setModel(updModel);
                   
                } catch (FileNotFoundException ex) {
                    Logger.getLogger(AnaSayfa.class.getName()).log(Level.SEVERE, null, ex);
                } catch (IOException ex) {
                    Logger.getLogger(AnaSayfa.class.getName()).log(Level.SEVERE, null, ex);
                } catch (EncryptedDocumentException ex) {
                    Logger.getLogger(AnaSayfa.class.getName()).log(Level.SEVERE, null, ex);
                }
      
       
        
        
        
    }//GEN-LAST:event_btn_updateActionPerformed

    private void txtSearchKeyTyped(java.awt.event.KeyEvent evt) {//GEN-FIRST:event_txtSearchKeyTyped

        TableRowSorter<DefaultTableModel> tr = new TableRowSorter<DefaultTableModel>((DefaultTableModel) jTable1.getModel());

            jTable1.setRowSorter(tr);
        
        tr.setRowFilter(RowFilter.regexFilter(txtSearch.getText().trim()));


        
    }//GEN-LAST:event_txtSearchKeyTyped

    
    private void jTable1MouseClicked(java.awt.event.MouseEvent evt) {//GEN-FIRST:event_jTable1MouseClicked

          updText.setText(jTable1.getValueAt(jTable1.getSelectedRow(), jTable1.getSelectedColumn()).toString());
        
    }//GEN-LAST:event_jTable1MouseClicked

    private void btn_excelActionPerformed(java.awt.event.ActionEvent evt) {//GEN-FIRST:event_btn_excelActionPerformed
        
        
        arr.clear();       
        
        jTable1.removeAll();
        model  = new DefaultTableModel();    
        jTable1.setModel(model);

        model.getDataVector().removeAllElements();
        model.fireTableDataChanged();
        
        JFileChooser dosyaSecici = new JFileChooser();
        int secim = dosyaSecici.showOpenDialog(this);
        
        if(secim == JFileChooser.APPROVE_OPTION){
            
            File dosya = dosyaSecici.getSelectedFile();
            dosya_yol.setText(dosya.getAbsolutePath());
            
        if(!dosya.getName().endsWith("xlsx")){
                        
            JOptionPane.showMessageDialog(this, "lutfen excel dosyası seçin","HATA", JOptionPane.ERROR_MESSAGE);
                
        }else
        {
              
            String excelFilePath= dosya_yol.getText();
		try
                {
                    
                    FileInputStream inputstream=new FileInputStream(excelFilePath);                        

                    XSSFWorkbook workbook=new XSSFWorkbook(inputstream);
                    XSSFSheet sheet=workbook.getSheetAt(0);	

                    int rows=sheet.getLastRowNum();
                    int cols=sheet.getRow(1).getLastCellNum();



                    for(int r=0;r<=rows;r++)
                    {
                            XSSFRow row=sheet.getRow(r); 

                            for(int c=0;c<cols;c++)
                            {

                            XSSFCell cell=row.getCell(c,org.apache.poi.ss.usermodel.Row.MissingCellPolicy.CREATE_NULL_AS_BLANK);


                                    if(r==0){
                                        model.addColumn(row.getCell(c));
                                    }

                                    CellType type = cell.getCellType();


                                    switch(type)
                                    {
                                    case STRING: arr.add(cell.getStringCellValue()); break;
                                    case NUMERIC: arr.add(cell.getNumericCellValue());break;
                                    case BOOLEAN: arr.add(cell.getBooleanCellValue()); break;
                                    case BLANK: arr.add(null); break;

                                    }

                            }
                    } 

                    for(int i = 0; i< arr.size() ; i=i+cols){

                            if(i==0){                     
                                continue;
                            }   
                             else{

                                  Vector vector = new Vector();

                            for(int k=0; k< cols; k++){

                               vector.add(arr.get(i+k));
                            }

                        model.addRow(vector);

                            }
                        }
                 } 
                catch (Exception x){
                    x.printStackTrace();
                }
                    
                }
        }

            
    }//GEN-LAST:event_btn_excelActionPerformed

    private void dosya_yolActionPerformed(java.awt.event.ActionEvent evt) {//GEN-FIRST:event_dosya_yolActionPerformed
        // TODO add your handling code here:
    }//GEN-LAST:event_dosya_yolActionPerformed

    
    public static void main(String args[]) {
        /* Set the Nimbus look and feel */
        //<editor-fold defaultstate="collapsed" desc=" Look and feel setting code (optional) ">
        /* If Nimbus (introduced in Java SE 6) is not available, stay with the default look and feel.
         * For details see http://download.oracle.com/javase/tutorial/uiswing/lookandfeel/plaf.html 
         */
        try {
            for (javax.swing.UIManager.LookAndFeelInfo info : javax.swing.UIManager.getInstalledLookAndFeels()) {
                if ("Nimbus".equals(info.getName())) {
                    javax.swing.UIManager.setLookAndFeel(info.getClassName());
                    break;
                }
            }
        } catch (ClassNotFoundException ex) {
            java.util.logging.Logger.getLogger(AnaSayfa.class.getName()).log(java.util.logging.Level.SEVERE, null, ex);
        } catch (InstantiationException ex) {
            java.util.logging.Logger.getLogger(AnaSayfa.class.getName()).log(java.util.logging.Level.SEVERE, null, ex);
        } catch (IllegalAccessException ex) {
            java.util.logging.Logger.getLogger(AnaSayfa.class.getName()).log(java.util.logging.Level.SEVERE, null, ex);
        } catch (javax.swing.UnsupportedLookAndFeelException ex) {
            java.util.logging.Logger.getLogger(AnaSayfa.class.getName()).log(java.util.logging.Level.SEVERE, null, ex);
        }
        //</editor-fold>

        /* Create and display the form */
        java.awt.EventQueue.invokeLater(new Runnable() {
            public void run() {
                try {
                    new AnaSayfa().setVisible(true);
                } catch (IOException ex) {
                    Logger.getLogger(AnaSayfa.class.getName()).log(Level.SEVERE, null, ex);
                }
            }
        });
    }

    // Variables declaration - do not modify//GEN-BEGIN:variables
    private javax.swing.JButton btn_excel;
    private javax.swing.JButton btn_update;
    private javax.swing.JTextField dosya_yol;
    private javax.swing.JLabel jLabel1;
    private javax.swing.JLabel jLabel2;
    private javax.swing.JScrollPane jScrollPane1;
    private javax.swing.JTable jTable1;
    private javax.swing.JTextField txtSearch;
    private javax.swing.JTextField updText;
    // End of variables declaration//GEN-END:variables
}
