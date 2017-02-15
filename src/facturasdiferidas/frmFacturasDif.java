package facturasdiferidas;

/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */
import com.toedter.calendar.JDateChooser;
import java.awt.Desktop;
import java.awt.Image;
import java.awt.Toolkit;
import java.awt.event.ActionEvent;
import java.awt.event.ActionListener;
import java.io.BufferedReader;
import java.io.File;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.FileReader;
import java.io.IOException;
import java.io.InputStream;
import java.net.URISyntaxException;
import java.net.URL;
import java.security.CodeSource;
import java.sql.Connection;
import java.sql.ResultSet;
import java.sql.SQLException;
import java.sql.Statement;
import java.text.SimpleDateFormat;
import java.util.Calendar;
import java.util.GregorianCalendar;
import java.util.Properties;
import java.util.logging.Level;
import java.util.logging.Logger;
import javax.swing.ImageIcon;
import javax.swing.JFileChooser;
import javax.swing.JOptionPane;
import javax.swing.Timer;
import javax.swing.filechooser.FileNameExtensionFilter;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.CreationHelper;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFFont;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

/**
 *
 * @author Usuario
 */
public class frmFacturasDif extends javax.swing.JFrame {

    
    Timer reloj;
    int segundos; // Una variable para manejar la cuenta regresiva.
    int copiasegundos; // Para recordar los segundos en caso de reiniciar la cuenta regresiva.
    int delay = 100;
    /**
     * Creates new form frmFacturasDif
     * @throws java.lang.ClassNotFoundException
     * @throws java.sql.SQLException
     */
    public frmFacturasDif() throws ClassNotFoundException, SQLException {
        initComponents();
          this.setLocationRelativeTo(null);
          
           URL pathIcon = this.getClass().getClassLoader().getResource("images/repmon.png");
            Toolkit kit = Toolkit.getDefaultToolkit();
            Image img = kit.createImage(pathIcon);
            this.setIconImage(img);
// Aprovechamos para ponerle el nombre de la aplicación
            this.setTitle("Reporte de Facturas diferidas");
//            FrmReporte.setDefaultLookAndFeelDecorated(true);
//            SubstanceLookAndFeel.setSkin("org.jvnet.substance.skin.CremeCoffeeSkin");
//            SubstanceLookAndFeel.setCurrentTheme("org.jvnet.substance.theme.SubstanceCremeTheme");
     
            
        //pathIcon = this.getClass().getClassLoader().getResource("images/DataBase.png");            
        ImageIcon imgicon=new ImageIcon("excelOri.png");
        btnExportar.setIcon(imgicon);
        
          
               
        grupo1.add(rbtAgencia);
        grupo1.add(rbtTodos);
        ConectarCombo();
         if (rbtAgencia.isSelected()){
            cboAgencias.setEnabled(true);
            op=true;
        }
        Calendar c2= new GregorianCalendar();
        dtcInicio.setCalendar(c2);
        dtpFin.setCalendar(c2);
        dtpCorte.setCalendar(c2);
         }
   
    ResultSet rsp;
    int last;
    private ResultSet FacturasDiferidas(String fecini,String fecfin,String feccorte, String agencia,Boolean opc) throws ClassNotFoundException, SQLException{
     
        Connection cn=clsConexion.Cadena();
        Statement St= cn.createStatement(ResultSet.TYPE_SCROLL_INSENSITIVE, ResultSet.CONCUR_READ_ONLY);
        String strop="";
        String stres="";
        String strini="DECLARE @fechaguia DATE SET @fechaguia='"+ feccorte +"' ";
         if (opc==true){
             strop=" and f6_ccodage IN ('" + Extraer(agencia) + "')  " +
            " and f6_dfecdoc between '" + fecini +"' and '"  + fecfin  + "' " +
            " order by 1,2,3,5,6";
        }
        else{
            strop=" and f6_dfecdoc between '" + fecini +"' and '"  + fecfin  + "' " +
            " order by 1,2,3,5,6";  
        }   
         
         stres=strini+leer()+strop;
         rsp=St.executeQuery(stres);
         
         rsp.last();
         last=rsp.getRow()+1;
         rsp.beforeFirst();
        
         return rsp;
        
    }

public String leer(){
    
  String ruta = System.getProperty("user.dir") + "\\script.sql";
  
  //Y generamos el objeto respecto a la ruta del archivo
  File archivo = new File(ruta);
  
  //Creamos la variable donde se almacenara la linea de texto leída
  String linea ="";
  String linea2="";
  
  //Y colocamos el try y el catch
  try {
   //Ahora creamos los objetos necesarios para leer el archivo.
   //Usamos este objeto que leera los bytes del archivo que ya creamos.
   FileReader lector = new FileReader(archivo);
   
   //Y creamos el buffer que guardara los bytes leidos pasandole al lector que lee los bytes.
   BufferedReader buff = new BufferedReader(lector);
   
   //Creamos el while, leyendo la linea en la condicion
   //Y se lee hasta que el valor leido sea nulo, es decir cuando ya no haya mas lineas
   while( ( linea = buff.readLine() ) != null ) {
    //Aqui dentro manejamos la linea
   //System.out.println(linea+"hola"); 
       linea2=linea2 +  linea + " ";
    //System.out.println(linea);
   }
   //System.out.println(linea+"hola");
   //cerramos los objetos
   buff.close();
   lector.close();
 
}       catch (FileNotFoundException ex) {
            Logger.getLogger(frmFacturasDif.class.getName()).log(Level.SEVERE, null, ex);
        } catch (IOException ex) {
            Logger.getLogger(frmFacturasDif.class.getName()).log(Level.SEVERE, null, ex);
   }
  return linea2;
}
    
    /**
     * This method is called from within the constructor to initialize the form.
     * WARNING: Do NOT modify this code. The content of this method is always
     * regenerated by the Form Editor.
     */
    @SuppressWarnings("unchecked")
    // <editor-fold defaultstate="collapsed" desc="Generated Code">//GEN-BEGIN:initComponents
    private void initComponents() {

        grupo1 = new javax.swing.ButtonGroup();
        btnExportar = new javax.swing.JButton();
        dtcInicio = new com.toedter.calendar.JDateChooser();
        dtpFin = new com.toedter.calendar.JDateChooser();
        jPanel1 = new javax.swing.JPanel();
        rbtAgencia = new javax.swing.JRadioButton();
        rbtTodos = new javax.swing.JRadioButton();
        cboAgencias = new javax.swing.JComboBox();
        jLabel1 = new javax.swing.JLabel();
        jLabel2 = new javax.swing.JLabel();
        jProgressBar1 = new javax.swing.JProgressBar();
        lblProgreso = new javax.swing.JLabel();
        dtpCorte = new com.toedter.calendar.JDateChooser();
        jLabel3 = new javax.swing.JLabel();

        setDefaultCloseOperation(javax.swing.WindowConstants.EXIT_ON_CLOSE);

        btnExportar.addActionListener(new java.awt.event.ActionListener() {
            public void actionPerformed(java.awt.event.ActionEvent evt) {
                btnExportarActionPerformed(evt);
            }
        });

        rbtAgencia.setSelected(true);
        rbtAgencia.setText("Agencia");
        rbtAgencia.addChangeListener(new javax.swing.event.ChangeListener() {
            public void stateChanged(javax.swing.event.ChangeEvent evt) {
                rbtAgenciaStateChanged(evt);
            }
        });

        rbtTodos.setText("Todos");
        rbtTodos.addChangeListener(new javax.swing.event.ChangeListener() {
            public void stateChanged(javax.swing.event.ChangeEvent evt) {
                rbtTodosStateChanged(evt);
            }
        });

        cboAgencias.setModel(new javax.swing.DefaultComboBoxModel(new String[] { "Item 1", "Item 2", "Item 3", "Item 4" }));
        cboAgencias.setEnabled(false);

        javax.swing.GroupLayout jPanel1Layout = new javax.swing.GroupLayout(jPanel1);
        jPanel1.setLayout(jPanel1Layout);
        jPanel1Layout.setHorizontalGroup(
            jPanel1Layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
            .addGroup(jPanel1Layout.createSequentialGroup()
                .addGap(19, 19, 19)
                .addGroup(jPanel1Layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
                    .addGroup(jPanel1Layout.createSequentialGroup()
                        .addComponent(rbtAgencia)
                        .addGap(28, 28, 28)
                        .addComponent(cboAgencias, javax.swing.GroupLayout.PREFERRED_SIZE, 188, javax.swing.GroupLayout.PREFERRED_SIZE))
                    .addComponent(rbtTodos))
                .addContainerGap(javax.swing.GroupLayout.DEFAULT_SIZE, Short.MAX_VALUE))
        );
        jPanel1Layout.setVerticalGroup(
            jPanel1Layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
            .addGroup(jPanel1Layout.createSequentialGroup()
                .addGap(20, 20, 20)
                .addGroup(jPanel1Layout.createParallelGroup(javax.swing.GroupLayout.Alignment.BASELINE)
                    .addComponent(rbtAgencia)
                    .addComponent(cboAgencias, javax.swing.GroupLayout.PREFERRED_SIZE, javax.swing.GroupLayout.DEFAULT_SIZE, javax.swing.GroupLayout.PREFERRED_SIZE))
                .addGap(18, 18, 18)
                .addComponent(rbtTodos)
                .addContainerGap(18, Short.MAX_VALUE))
        );

        jLabel1.setText("Inicio");

        jLabel2.setText("Final");

        lblProgreso.setText(".");

        jLabel3.setText("Corte");

        javax.swing.GroupLayout layout = new javax.swing.GroupLayout(getContentPane());
        getContentPane().setLayout(layout);
        layout.setHorizontalGroup(
            layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
            .addGroup(layout.createSequentialGroup()
                .addContainerGap()
                .addGroup(layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
                    .addComponent(jPanel1, javax.swing.GroupLayout.PREFERRED_SIZE, javax.swing.GroupLayout.DEFAULT_SIZE, javax.swing.GroupLayout.PREFERRED_SIZE)
                    .addComponent(jLabel1)
                    .addGroup(layout.createSequentialGroup()
                        .addGroup(layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
                            .addComponent(dtpFin, javax.swing.GroupLayout.PREFERRED_SIZE, 147, javax.swing.GroupLayout.PREFERRED_SIZE)
                            .addComponent(jLabel2)
                            .addComponent(dtcInicio, javax.swing.GroupLayout.PREFERRED_SIZE, 147, javax.swing.GroupLayout.PREFERRED_SIZE))
                        .addGap(26, 26, 26)
                        .addComponent(btnExportar, javax.swing.GroupLayout.PREFERRED_SIZE, 108, javax.swing.GroupLayout.PREFERRED_SIZE))
                    .addGroup(layout.createSequentialGroup()
                        .addGroup(layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
                            .addComponent(dtpCorte, javax.swing.GroupLayout.PREFERRED_SIZE, 147, javax.swing.GroupLayout.PREFERRED_SIZE)
                            .addComponent(jLabel3))
                        .addPreferredGap(javax.swing.LayoutStyle.ComponentPlacement.RELATED)
                        .addGroup(layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING, false)
                            .addGroup(layout.createSequentialGroup()
                                .addGap(25, 25, 25)
                                .addComponent(lblProgreso, javax.swing.GroupLayout.DEFAULT_SIZE, javax.swing.GroupLayout.DEFAULT_SIZE, Short.MAX_VALUE))
                            .addComponent(jProgressBar1, javax.swing.GroupLayout.PREFERRED_SIZE, javax.swing.GroupLayout.DEFAULT_SIZE, javax.swing.GroupLayout.PREFERRED_SIZE))))
                .addContainerGap(javax.swing.GroupLayout.DEFAULT_SIZE, Short.MAX_VALUE))
        );
        layout.setVerticalGroup(
            layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
            .addGroup(layout.createSequentialGroup()
                .addGap(8, 8, 8)
                .addComponent(jPanel1, javax.swing.GroupLayout.PREFERRED_SIZE, javax.swing.GroupLayout.DEFAULT_SIZE, javax.swing.GroupLayout.PREFERRED_SIZE)
                .addGap(18, 18, 18)
                .addComponent(jLabel1)
                .addPreferredGap(javax.swing.LayoutStyle.ComponentPlacement.RELATED)
                .addGroup(layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
                    .addComponent(btnExportar, javax.swing.GroupLayout.PREFERRED_SIZE, 67, javax.swing.GroupLayout.PREFERRED_SIZE)
                    .addGroup(layout.createParallelGroup(javax.swing.GroupLayout.Alignment.TRAILING)
                        .addGroup(layout.createSequentialGroup()
                            .addComponent(jProgressBar1, javax.swing.GroupLayout.PREFERRED_SIZE, javax.swing.GroupLayout.DEFAULT_SIZE, javax.swing.GroupLayout.PREFERRED_SIZE)
                            .addPreferredGap(javax.swing.LayoutStyle.ComponentPlacement.RELATED)
                            .addComponent(lblProgreso))
                        .addGroup(layout.createSequentialGroup()
                            .addComponent(dtcInicio, javax.swing.GroupLayout.PREFERRED_SIZE, javax.swing.GroupLayout.DEFAULT_SIZE, javax.swing.GroupLayout.PREFERRED_SIZE)
                            .addPreferredGap(javax.swing.LayoutStyle.ComponentPlacement.RELATED)
                            .addComponent(jLabel2, javax.swing.GroupLayout.PREFERRED_SIZE, 25, javax.swing.GroupLayout.PREFERRED_SIZE)
                            .addPreferredGap(javax.swing.LayoutStyle.ComponentPlacement.RELATED)
                            .addComponent(dtpFin, javax.swing.GroupLayout.PREFERRED_SIZE, javax.swing.GroupLayout.DEFAULT_SIZE, javax.swing.GroupLayout.PREFERRED_SIZE)
                            .addPreferredGap(javax.swing.LayoutStyle.ComponentPlacement.RELATED)
                            .addComponent(jLabel3, javax.swing.GroupLayout.PREFERRED_SIZE, 25, javax.swing.GroupLayout.PREFERRED_SIZE)
                            .addPreferredGap(javax.swing.LayoutStyle.ComponentPlacement.RELATED)
                            .addComponent(dtpCorte, javax.swing.GroupLayout.PREFERRED_SIZE, javax.swing.GroupLayout.DEFAULT_SIZE, javax.swing.GroupLayout.PREFERRED_SIZE))))
                .addContainerGap(javax.swing.GroupLayout.DEFAULT_SIZE, Short.MAX_VALUE))
        );

        pack();
    }// </editor-fold>//GEN-END:initComponents

    
    
    private class EjecutarTarea2 implements ActionListener{

        @Override
        public void actionPerformed(ActionEvent ae) {
        try {                  
        int n=jProgressBar1.getValue();
          if (rsp.next()==true){
              jProgressBar1.setValue(1+Math.round(100*f/last-1));
              int pr=1+Math.round(100*f/last-1);
                            
              lblProgreso.setText(Integer.toString(pr)+" %");
              if(f==last-1){
              jProgressBar1.setValue(100);    
              lblProgreso.setText(100 + " %");
              }
              
              Row fila=hoja.createRow(f);
              fila=hoja.createRow(f);
            Cell celda1=fila.createCell(0);
            celda1.setCellValue(rsp.getString("TIENDA"));
            
            celda1=fila.createCell(1);
            celda1.setCellValue(rsp.getString("TIPO_DOC"));
            
            celda1=fila.createCell(2);
            celda1.setCellValue(rsp.getString("DOC_DIFERIDO"));
            
            celda1=fila.createCell(3);
            celda1.setCellValue(rsp.getDate("FECHA_EMI"));
            celda1.setCellStyle(cellStyle);
            
            celda1=fila.createCell(4);
            celda1.setCellValue(rsp.getString("CLIENTE"));
            
            celda1=fila.createCell(5);
            celda1.setCellValue(rsp.getString("NOMBRE"));
                                   
            celda1=fila.createCell(6);
            celda1.setCellValue(rsp.getString("COD_PROD"));
            
            celda1=fila.createCell(7);
            celda1.setCellValue(rsp.getString("PRODUCTO"));
            
            celda1=fila.createCell(8);
            celda1.setCellValue(rsp.getString("MONEDA"));
            
            celda1=fila.createCell(9);
            celda1.setCellValue(rsp.getDouble("PRECIO_UNITARIO"));
            
            celda1=fila.createCell(10);
            celda1.setCellValue(rsp.getDouble("TOTAL"));
            
            celda1=fila.createCell(11);
            celda1.setCellValue(rsp.getDouble("CANT_FACTURADA"));
            
            celda1=fila.createCell(12);
            celda1.setCellValue(rsp.getDouble("SALDO_TOTAL_X_ATENDER"));
            
            celda1=fila.createCell(13);
            celda1.setCellValue(rsp.getDouble("CANT_ATENDIDA"));
            
            celda1=fila.createCell(14);
            celda1.setCellValue(rsp.getString("TIPO_DOC_A"));
            
            celda1=fila.createCell(15);
            celda1.setCellValue(rsp.getString("DOCUMENTO_ATIENDE"));
            
            celda1=fila.createCell(16);
            celda1.setCellValue(rsp.getDate("FECHA_GUIA"));
            celda1.setCellStyle(cellStyle);
            
            celda1=fila.createCell(17);
            celda1.setCellValue(rsp.getString("TIENDA_ATIENDE"));
            
            celda1=fila.createCell(18);
            celda1.setCellValue(rsp.getString("ALMACEN_ATIENDE"));
            
            
        f=f+1;         
              
            }
     
            else
             {
             reloj.stop();
             hoja.autoSizeColumn((short)0);
            hoja.autoSizeColumn((short)1);
            hoja.autoSizeColumn((short)2);
            hoja.autoSizeColumn((short)3);
            hoja.autoSizeColumn((short)4);
            hoja.autoSizeColumn((short)5);
            hoja.autoSizeColumn((short)6);
            hoja.autoSizeColumn((short)7);
            hoja.autoSizeColumn((short)8);
            hoja.autoSizeColumn((short)9);
            hoja.autoSizeColumn((short)10);
            hoja.autoSizeColumn((short)11);
            hoja.autoSizeColumn((short)12);
            hoja.autoSizeColumn((short)13);
            hoja.autoSizeColumn((short)14);
            hoja.autoSizeColumn((short)15);
            hoja.autoSizeColumn((short)16);
            hoja.autoSizeColumn((short)17);
            hoja.autoSizeColumn((short)18);
                    
             libro.write(archivo);
        /*Cerramos el flujo de datos*/
                      archivo.close();
        /*Y abrimos el archivo con la clase Desktop*/
                      Desktop.getDesktop().open(archivoXLS);
                      
                      JOptionPane.showMessageDialog(null,"Generación Completa\n Se creo el archivo en "+rutaArchivo);
                    }
           
                   } catch (IOException | SQLException ex) {
                           JOptionPane.showMessageDialog(null,ex,"Advertencia",JOptionPane.WARNING_MESSAGE);
                      }
            
        }
        
    }
    
    
    
    private void ConectarCombo() throws ClassNotFoundException, SQLException {
    
           cboAgencias.removeAllItems();
        
        try (
        
        Connection cn=clsConexion.Cadena();
        Statement St=cn.createStatement(ResultSet.TYPE_SCROLL_INSENSITIVE, ResultSet.CONCUR_READ_ONLY);
        )
        {   
            ResultSet rs=St.executeQuery("SELECT LTRIM(RTRIM(AG_CCODAGE +' :' +AG_CDESCRI)) AG_CDESCRI FROM FT0001AGEN");
        //rs.first();
            while (rs.next())
            {
            cboAgencias.addItem(rs.getString("AG_CDESCRI"));
            //JOptionPane.showMessageDialog(this,Extraer(rs.getString("AG_CDESCRI")));
            } 
        }      
            
}
    
            Workbook libro;
            File archivoXLS;
            FileOutputStream archivo;
            Sheet hoja;
            String rutaArchivo;
            Font fuente,fuente2,fuente3;
            int f=0;
            CellStyle cellStyle;
    
    private void ExportarExcel(ResultSet rs) throws IOException, SQLException{
        
    JFileChooser file=new JFileChooser();
    FileNameExtensionFilter filtro = new FileNameExtensionFilter("Excel(*.XLSX)", "xlsx"); 
    file.setFileFilter(filtro);
    int seleccion=file.showSaveDialog(this);
    
       
    if (seleccion==JFileChooser.CANCEL_OPTION){
        return;
    }
    File guarda =file.getSelectedFile();
    
       
    
    //file.set
    
       if(guarda !=null)
        {
           // file.setFileFilter(filtro);
       /*guardamos el archivo y le damos el formato directamente,
        * si queremos que se guarde en formato doc lo definimos como .doc*/
            rutaArchivo = guarda.getAbsolutePath()+".xlsx";
    //    JOptionPane.showMessageDialog(null,
    //         "Se creo el archivo... Generando información",
    //             "Información en\n"+rutaArchivo,JOptionPane.INFORMATION_MESSAGE);

       }
       else
       {
        //   file.setFileFilter(filtro);
               rutaArchivo = System.getProperty("user.home")+"/Kardex.xlsx";
    //    JOptionPane.showMessageDialog(null,
    //         "Se creo el archivo... Generando información",
    //             "Información en\n"+rutaArchivo,JOptionPane.INFORMATION_MESSAGE);
       }
 /*Se crea el objeto de tipo File con la ruta del archivo*/
        archivoXLS = new File(rutaArchivo);
        /*Si el archivo existe se elimina*/
        if(archivoXLS.exists()) archivoXLS.delete();
//        try {
            /*Se crea el archivo*/
            //archivoXLS.
            archivoXLS.createNewFile();
            libro = new XSSFWorkbook(); 
            archivo = new FileOutputStream(archivoXLS);       
            //archivo.close();
                   
        CreationHelper createhelper=libro.getCreationHelper();
         cellStyle=libro.createCellStyle();
        
        cellStyle.setDataFormat(createhelper.createDataFormat().getFormat("dd/mm/yyyy"));
        
        fuente =libro.createFont();
        fuente.setFontHeightInPoints((short)10);
        fuente.setFontName("Calibri");
        fuente.setBoldweight(XSSFFont.BOLDWEIGHT_BOLD);
        fuente.setBold(true);
        
        fuente3 =libro.createFont();
        fuente3.setFontHeightInPoints((short)8);
        fuente3.setFontName("Calibri");
        fuente3.setBoldweight(XSSFFont.BOLDWEIGHT_BOLD);
        
        
        fuente2 =libro.createFont();
        fuente2.setFontHeightInPoints((short)9);
        fuente2.setFontName("Calibri");
        //fuente.setBoldweight(XSSFFont.BOLDWEIGHT_BOLD);
        
        CellStyle cellStyle2=libro.createCellStyle();
        cellStyle2.setAlignment(XSSFCellStyle. ALIGN_CENTER);
        cellStyle2.setVerticalAlignment(XSSFCellStyle.VERTICAL_TOP);
        cellStyle2.setFont(fuente);
        cellStyle2.setWrapText(true);
        
        CellStyle cellStyle3=libro.createCellStyle();
        cellStyle3.setAlignment(XSSFCellStyle. ALIGN_LEFT);
        cellStyle3.setVerticalAlignment(XSSFCellStyle.VERTICAL_TOP);
        cellStyle3.setFont(fuente);
        
        CellStyle cellStyle4=libro.createCellStyle();
//        cellStyle4.setAlignment(XSSFCellStyle. ALIGN_LEFT);
//        cellStyle4.setVerticalAlignment(XSSFCellStyle.VERTICAL_TOP);
        cellStyle4.setFont(fuente);
        
        hoja = libro.createSheet("FacturasDiferidas");
        //hoja.
        
        hoja.setDefaultRowHeightInPoints(12);
       
        Row fila=hoja.createRow(0);
        
        Cell celda=fila.createCell(0);
        celda.setCellValue("TIENDA");
        celda.setCellStyle(cellStyle4);
                        
        celda=fila.createCell(1);
        celda.setCellValue("TIPO_DOC");
        celda.setCellStyle(cellStyle4);
        celda=fila.createCell(2);
        celda.setCellValue("DOC_DIFERIDO");
        celda.setCellStyle(cellStyle4);
        celda=fila.createCell(3);
        celda.setCellValue("FECHA_EMI");
        celda.setCellStyle(cellStyle4);
        celda=fila.createCell(4);
        celda.setCellValue("CLIENTE");
        celda.setCellStyle(cellStyle4);
        celda=fila.createCell(5);
        celda.setCellValue("NOMBRE");
        celda.setCellStyle(cellStyle4);
        celda=fila.createCell(6);
        celda.setCellValue("COD_PROD");
        celda.setCellStyle(cellStyle4);
        celda=fila.createCell(7);
        celda.setCellValue("PRODUCTO");
        celda.setCellStyle(cellStyle4);
        celda=fila.createCell(8);
        celda.setCellValue("MONEDA");
        celda.setCellStyle(cellStyle4);
        celda=fila.createCell(9);
        celda.setCellValue("PRECIO_UNITARIO");
        celda.setCellStyle(cellStyle4);
        celda=fila.createCell(10);
        celda.setCellValue("TOTAL");
        celda.setCellStyle(cellStyle4);
        celda=fila.createCell(11);
        celda.setCellValue("CANT_FACTURADA");
        celda.setCellStyle(cellStyle4);
        celda=fila.createCell(12);
        celda.setCellValue("SALDO_TOTAL_X_ATENDER");
        celda.setCellStyle(cellStyle4);
        celda=fila.createCell(13);
        celda.setCellValue("CANTIDAD_ATENDIDA");
        celda.setCellStyle(cellStyle4);
        celda=fila.createCell(14);
        celda.setCellValue("TIPO_DOC_A");
        celda.setCellStyle(cellStyle4);
        celda=fila.createCell(15);
        celda.setCellValue("DOCUMENTO_ATIENDE");
        celda.setCellStyle(cellStyle4);
        celda=fila.createCell(16);
        celda.setCellValue("FECHA_GUIA");
        celda.setCellStyle(cellStyle4);
        celda=fila.createCell(17);
        celda.setCellValue("TIENDA_ATIENDE");
        celda.setCellStyle(cellStyle4);
        celda=fila.createCell(18);
        celda.setCellValue("ALMACEN_ATIENDE");
        celda.setCellStyle(cellStyle4);
        f=1;     
        reloj = new Timer(delay,new EjecutarTarea2() );
        reloj.start();
        //           
//        
//        while (rs.next()){
//                       
//        }
        
    }
    private String ConvertDate(JDateChooser jd) {
     SimpleDateFormat formato;
        formato = new SimpleDateFormat("yyyy-MM-dd");
        if (jd.getDate()!=null){
        return formato.format(jd.getDate());
    }
    else
    {
        return null;
    }
  }
    private String Extraer(String cadena){
        String cadenaux;
        
        int pos=cadena.indexOf(":");
        cadenaux=cadena.substring(0,pos-1);
        return cadenaux;
    }
    
    private void btnExportarActionPerformed(java.awt.event.ActionEvent evt) {//GEN-FIRST:event_btnExportarActionPerformed
        try {
            // TODO add your handling code here:
            
        Connection cn=clsConexion.Cadena();
        Statement St= cn.createStatement();
        FacturasDiferidas(ConvertDate(dtcInicio), ConvertDate(dtpFin),ConvertDate(dtpCorte), (String) cboAgencias.getSelectedItem(),op);
        if(!rsp.next()){
             JOptionPane.showMessageDialog(this,"No hay datos con los criterios especificados");   
         }
        else {
            rsp.beforeFirst();
            ExportarExcel(rsp);
        }
            
        } catch (ClassNotFoundException | SQLException | IOException ex) {
            Logger.getLogger(frmFacturasDif.class.getName()).log(Level.SEVERE, null, ex);
        }
        
    }//GEN-LAST:event_btnExportarActionPerformed

    boolean op=false;
    
    private void rbtAgenciaStateChanged(javax.swing.event.ChangeEvent evt) {//GEN-FIRST:event_rbtAgenciaStateChanged
        // TODO add your handling code here:
        if (rbtAgencia.isSelected()){
            cboAgencias.setEnabled(true);
            op=true;
        }
    }//GEN-LAST:event_rbtAgenciaStateChanged

    private void rbtTodosStateChanged(javax.swing.event.ChangeEvent evt) {//GEN-FIRST:event_rbtTodosStateChanged
        // TODO add your handling code here:
         if (rbtTodos.isSelected()){
            cboAgencias.setEnabled(false);
            op=false;
        }
    }//GEN-LAST:event_rbtTodosStateChanged

    /**
     * @param args the command line arguments
     */
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
        } catch (ClassNotFoundException | InstantiationException | IllegalAccessException | javax.swing.UnsupportedLookAndFeelException ex) {
            java.util.logging.Logger.getLogger(frmFacturasDif.class.getName()).log(java.util.logging.Level.SEVERE, null, ex);
        }
        //</editor-fold>
        
        //</editor-fold>

        /* Create and display the form */
        java.awt.EventQueue.invokeLater(() -> {
            try {
                new frmFacturasDif().setVisible(true);
            } catch (ClassNotFoundException | SQLException ex) {
                Logger.getLogger(frmFacturasDif.class.getName()).log(Level.SEVERE, null, ex);
            }
        });
    }
    
    

    // Variables declaration - do not modify//GEN-BEGIN:variables
    private javax.swing.JButton btnExportar;
    private javax.swing.JComboBox cboAgencias;
    private com.toedter.calendar.JDateChooser dtcInicio;
    private com.toedter.calendar.JDateChooser dtpCorte;
    private com.toedter.calendar.JDateChooser dtpFin;
    private javax.swing.ButtonGroup grupo1;
    private javax.swing.JLabel jLabel1;
    private javax.swing.JLabel jLabel2;
    private javax.swing.JLabel jLabel3;
    private javax.swing.JPanel jPanel1;
    private javax.swing.JProgressBar jProgressBar1;
    private javax.swing.JLabel lblProgreso;
    private javax.swing.JRadioButton rbtAgencia;
    private javax.swing.JRadioButton rbtTodos;
    // End of variables declaration//GEN-END:variables
}
