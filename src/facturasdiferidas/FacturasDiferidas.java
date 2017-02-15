/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */
package facturasdiferidas;

import java.io.IOException;
import java.io.InputStream;
import java.util.Properties;

/**
 *
 * @author Usuario
 */
public class FacturasDiferidas {
        
      
    /**
     * @param args the command line arguments
     * @throws java.io.IOException
     */
    public static void main(String[] args) throws IOException {

        // TODO code application logic here
        
         Properties props = new Properties();
         InputStream in = null;
        
            in = FacturasDiferidas.class.getResourceAsStream("/config.properties");
            props.load(in);
            in.close();
            
            String base1 = props.getProperty("base");
            String pass1 = props.getProperty("pass");
            String user1 = props.getProperty("user");
            String server1 = props.getProperty("server");
             
        System.out.println(base1);
        System.out.println(pass1);
        System.out.println(user1);
        System.out.println(server1);
                        
    }
    
}
