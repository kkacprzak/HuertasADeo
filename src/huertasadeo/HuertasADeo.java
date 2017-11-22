/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */
package huertasadeo;

import huertasadeo.Frames.MainFrame;
import java.awt.BorderLayout;
import java.io.File;
import javax.swing.JFrame;

/**
 *
 * @author krzysiek
 */
public class HuertasADeo {

    public static File mainFile = null;
    /**
     * @param args the command line arguments
     */
    public static void main(String[] args) {
        MainFrame mainFrame = new MainFrame();
        mainFrame.setDefaultCloseOperation(JFrame.EXIT_ON_CLOSE);
        mainFrame.pack();
        mainFrame.setExtendedState(mainFrame.getExtendedState() | JFrame.MAXIMIZED_BOTH);
        mainFrame.setVisible(true);
    }
    
}
