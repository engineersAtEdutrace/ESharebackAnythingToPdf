/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */


import com.artofsolving.jodconverter.DocumentConverter;
import com.artofsolving.jodconverter.openoffice.connection.OpenOfficeConnection;
import com.artofsolving.jodconverter.openoffice.connection.SocketOpenOfficeConnection;
import com.artofsolving.jodconverter.openoffice.converter.OpenOfficeDocumentConverter;
import java.io.File;
import java.io.IOException;
import java.net.ConnectException;
import java.net.InetAddress;
import java.net.Socket;
import java.net.UnknownHostException;
import java.util.logging.Level;
import java.util.logging.Logger;

/**
 *
 * @author Anand Singh
 * 
 * 
 */
public class AnythingToPdf {
    public static void main(String[] args)throws UnknownHostException, IOException {
        
        try{
            File inputFile = new File("D:\\Java_Workspace\\JodConverter\\Pptx\\sample_1.pptx");
            File outputFile = new File("D:\\Java_Workspace\\JodConverter\\Pptx\\sample_1_pptx.pdf");

            // connect to an OpenOffice.org instance running on port 8100
            OpenOfficeConnection connection = new SocketOpenOfficeConnection(8100);
            connection.connect();

            // convert
            DocumentConverter converter = new OpenOfficeDocumentConverter(connection);
            converter.convert(inputFile, outputFile);

            // close the connection
            connection.disconnect();
//        }else
//        {
//            System.out.println("Not connected to Openoffice, please connect");
//        }
            
        }catch(ConnectException e)
        {
            String err = "*** OpenOffice service in not running please run the following command ***";
            String cmd="\tcommand : soffice -headless -accept=\"socket,host=127.0.0.1,port=8100;urp;\" -nofirststartwizard\n\n";
            Logger.getLogger(AnythingToPdf.class.getName()).log(Level.SEVERE, err+"\n"+cmd);
            //System.out.println("OpenOffice is not connected,please connect");
        }
        
    }
    
}
