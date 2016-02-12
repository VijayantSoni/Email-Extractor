/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */
package cv_extractor;

import java.awt.BorderLayout;
import java.awt.Color;
import java.awt.GridBagLayout;
import java.awt.GridLayout;
import java.awt.HeadlessException;
import java.awt.Insets;
import java.awt.event.ActionEvent;
import java.awt.event.ActionListener;
import java.io.File;
import java.util.HashSet;
import java.util.Iterator;
import java.util.logging.Level;
import java.util.logging.Logger;
import javax.swing.JButton;
import javax.swing.JFileChooser;
import javax.swing.JFrame;
import javax.swing.JOptionPane;
import javax.swing.JPanel;
import javax.swing.JScrollPane;
import javax.swing.JTextArea;
import javax.swing.JTextField;
import javax.swing.SwingUtilities;
import javax.swing.UIManager;
import javax.swing.UnsupportedLookAndFeelException;
import javax.swing.filechooser.FileFilter;

/**
 *
 * @author Vijayant Soni
 */

public class CV_Extractor extends JFrame implements ActionListener
{
    static JButton select, generate, clear;
    
    static JTextField cvFileName;
    
    static JTextArea cvList;
    
    static JPanel pageEnd, pageStart;
    
    static JScrollPane cvListScrollPane;
    
    private CV_Extractor() throws HeadlessException 
    {
        setSize(500,500);
        setBackground(Color.yellow);
        
        setDefaultCloseOperation(JFrame.EXIT_ON_CLOSE);
        
        pageEnd = new JPanel();
        pageStart = new JPanel();
        
        select = new JButton("Select");
        select.addActionListener(this);
        select.setSize(100,100);
        
        generate = new JButton("Generate");
        generate.addActionListener(this);
        generate.setSize(100,100);
        generate.setForeground(Color.blue);
        
        clear = new JButton("Clear");
        clear.addActionListener(this);
        clear.setSize(100,100);
        clear.setForeground(Color.red);
        
        cvList = new JTextArea("No E-mails right now",20,20);
        cvList.setSize(300,300);
        cvList.setMargin(new Insets(5,5,5,5));
        cvList.setEditable(false);
        
        cvFileName = new JTextField();
        cvFileName.setSize(300,300);
        
        cvListScrollPane = new JScrollPane(cvList);
        
        pageEnd.add(select);
        pageEnd.add(generate);
        pageEnd.add(clear);
        
        pageStart.add(cvListScrollPane);
        
        add(pageEnd,BorderLayout.PAGE_END);
        add(pageStart,BorderLayout.PAGE_START);
        
        setVisible(true);        
    }
    
    public void actionPerformed(ActionEvent ae)
    {
        if(ae.getActionCommand().equals("Select"))
        {
            openFileChooserToSelectCV();
        }        
        else if(ae.getActionCommand().equals("Generate"))
        {
            if((cvList.getText().equals("No E-mails right now")))
            {
                JOptionPane.showMessageDialog(null,"No Emails !");
            }
            else
            {                
                openFileChooserToSelectDirectory();
            }
        }
         else if(ae.getActionCommand().equals("Clear"))
        {
            cvList.setText("No E-mails right now");
        }
    }
        
    public static void main(String[] args) 
    {
        SwingUtilities.invokeLater(new Runnable() {
            public void run()
            {
                try {
                    //Turn off metal's use of bold fonts
                    UIManager.put("swing.boldMetal", Boolean.FALSE);
                    UIManager.setLookAndFeel(UIManager.getSystemLookAndFeelClassName());
                    
                } 
                catch (ClassNotFoundException ex) {
                    Logger.getLogger(CV_Extractor.class.getName()).log(Level.SEVERE, null, ex);
                } 
                catch (InstantiationException ex) {
                    Logger.getLogger(CV_Extractor.class.getName()).log(Level.SEVERE, null, ex);
                } 
                catch (IllegalAccessException ex) {
                    Logger.getLogger(CV_Extractor.class.getName()).log(Level.SEVERE, null, ex);
                } 
                catch (UnsupportedLookAndFeelException ex) {
                    Logger.getLogger(CV_Extractor.class.getName()).log(Level.SEVERE, null, ex);
                }
                
                finally
                {
                    new CV_Extractor();
                }
            }
        });
    }
    
    private void openFileChooserToSelectCV()
    {
        //Filter to let user only select doc or docx file
        FileFilter filter = new FileFilter() 
        {
            @Override
            public boolean accept(File f) 
            {
                if (f.isDirectory())
                {
                  return true;
                }

            String fileName = f.getName().toLowerCase();
        
             if (fileName.endsWith(".doc") || fileName.endsWith(".docx")) 
             {
                 return true;
             }
             
             return false; // Reject any other files
            }
            
            @Override
            public String getDescription() 
            {
                return "Word Document";
            }
        };
            
        JFileChooser chooser = new JFileChooser("S://");
        
        chooser.setFileFilter(filter);
        
        chooser.setMultiSelectionEnabled(true);
        
        chooser.setDialogTitle("Select CV");
        
        chooser.setApproveButtonText("Select File");
        
        int returnVal = chooser.showOpenDialog(this);
    
         if(returnVal == JFileChooser.APPROVE_OPTION) 
        {
            File files[] = chooser.getSelectedFiles();
            
            for(int i=0;i<files.length;i++)
            {
                System.out.println(files[i].getName() + "\n");
            }
            
            new DocReader(files);
            
            DocReader.readDocxFile();
        }   
    }
    
    private void openFileChooserToSelectDirectory()
    {
        JFileChooser chooser = new JFileChooser();
        
        chooser.setFileSelectionMode(JFileChooser.DIRECTORIES_ONLY);
        
        chooser.setDialogTitle("Select folder to save Excel file");
        
        int option = chooser.showSaveDialog(this);
        
        try
        {
            if (option == JFileChooser.APPROVE_OPTION)
            {
                File file = chooser.getSelectedFile();
                
                System.out.println(file.getCanonicalPath());                             
                            
                DocReader.generateExcel(file.getCanonicalPath());
            }
        }
        catch(Exception e)
        {
            e.printStackTrace();
        }
        
    }
    
    
    protected static void showInTetxArea(HashSet<String> hashSet)
    {
        cvList.setText("");
        
        Iterator <String> iterator = hashSet.iterator();
        
        while(iterator.hasNext())
        {
            cvList.append(iterator.next()+"\n");
        }
        
    }
       
}