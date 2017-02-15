/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */
package facturasdiferidas;


/**
 *
 * @author Usuario
 */
import java.io.BufferedReader;
import java.io.File;
import java.io.FileNotFoundException;
import java.io.FileReader;
import java.io.IOException;
import java.io.InputStream;
import java.net.URISyntaxException;
import java.sql.Connection;
import java.sql.DriverManager;
import java.sql.SQLException;
import java.util.Properties;
import java.util.logging.Level;
import java.util.logging.Logger;
import javax.swing.JOptionPane;

public class clsConexion {
    
    static String servidor;
    static String puerto="1433";
    static String user;
    static String password;
    static String base;
    static String driver="com.microsoft.sqlserver.jdbc.SQLServerDriver";
    
      public static Connection Cadena() throws ClassNotFoundException, SQLException {
                    /*Class.forName("com.microsoft.sqlserver.jdbc.SQLServerDriver");*/
          Properties props = new Properties();
         InputStream in = null;
        
           // in = FacturasDiferidas.class.getResourceAsStream("/config.properties");
        try {
            props=obtenerArchivoPropiedades("config.properties");
        //    props.load(in);
         //in.close();
            base= props.getProperty("base");
            password= props.getProperty("pass");
            user = props.getProperty("user");
            servidor= props.getProperty("server");
                        
        } catch (IOException | URISyntaxException ex) {
            Logger.getLogger(frmFacturasDif.class.getName()).log(Level.SEVERE, null, ex);
        }                 
          
          
            Class.forName(driver);
            String url="jdbc:sqlserver://"+servidor+":"+puerto+";"+"databasename="+base+";user="+user+";password="+password+";";
            Connection cnn=DriverManager.getConnection(url);
            return cnn;                           
      
    }
    
    public static Properties obtenerArchivoPropiedades(String arc) throws FileNotFoundException, URISyntaxException {
    Properties prop = null;
    try {
       // CodeSource codeSource = FacturasDiferidas.class.getProtectionDomain().getCodeSource();
        File jarFile = new File( System.getProperty("user.dir"));
         
            File propFile = new File(jarFile, arc);
            prop = new Properties();
            prop.load(new BufferedReader(new FileReader(propFile.getAbsoluteFile())));
     
    } catch (IOException ex) {
        Logger.getLogger(FacturasDiferidas.class.getName()).log(Level.SEVERE, null, ex);
    }
    return prop;
}
      
}
