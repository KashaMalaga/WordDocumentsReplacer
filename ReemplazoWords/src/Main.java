import fr.opensagres.poi.xwpf.converter.pdf.PdfConverter;
import fr.opensagres.poi.xwpf.converter.pdf.PdfOptions;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.docx4j.openpackaging.packages.WordprocessingMLPackage;
import org.docx4j.wml.Text;

import javax.swing.*;
import javax.xml.bind.JAXBElement;
import java.awt.*;
import java.awt.event.ActionEvent;
import java.awt.event.ActionListener;
import java.io.*;
import java.nio.file.Files;
import java.text.SimpleDateFormat;
import java.time.Clock;
import java.time.LocalDate;
import java.time.format.DateTimeFormatter;
import java.util.Date;
import java.util.List;
import java.util.Locale;
import java.util.Objects;
import java.util.logging.*;

import static java.nio.file.Files.walk;
import static javax.swing.SwingUtilities.*;

@SuppressWarnings("resource")
public class Main implements ActionListener {
    private static final String XPATH_TO_SELECT_TEXT_NODES = "//w:t";
    static final Logger logger = Logger.getLogger(Main.class.getName());
    private static final JTextArea errorLog =new JTextArea(15,75);
    private static  JFrame frame;
    private static final int dialogButton = JOptionPane.YES_NO_OPTION;
    static int counter = 0;
    static File folder;
    public static void main(String[] args) {
        try {
            SimpleDateFormat dateFormat = new SimpleDateFormat("yyyy-MM-dd_HH-mm-ss");
            String fileName = String.format("logfile_%s.log", dateFormat.format(new Date()));
            FileHandler fileHandler = new FileHandler(fileName, true);
            fileHandler.setFormatter(new SimpleFormatter());
            logger.addHandler(fileHandler);

        } catch (IOException e) {
            e.printStackTrace();
        }


        frame = new JFrame("Word Docxs Updater/Replacer + PDF Converter");
        frame.setLayout(new FlowLayout());
        frame.pack();
        frame.setLocationRelativeTo(null);
        frame.setSize(new Dimension(930, 380));
        frame.setDefaultCloseOperation(JFrame.EXIT_ON_CLOSE);
        Dimension dim = Toolkit.getDefaultToolkit().getScreenSize();
        frame.setLocation(dim.width/2-frame.getSize().width/2, dim.height/2-frame.getSize().height/2);
        JLabel searchLabel = new JLabel("Word to search:");
        JTextField searchField = new JTextField(17);
        searchField.setText("NOMINA DEL MES DE");

        JLabel replaceLabel = new JLabel("Word to replace:");
        JTextField replaceField = new JTextField(17);
        Locale spanishLocale=new Locale("es", "ES");
        Clock cl = Clock.systemUTC();
        LocalDate localDate = LocalDate.now(cl).plusMonths(1);
        String proximoMesSpanish=localDate.format(DateTimeFormatter.ofPattern("MMMM",spanishLocale));
        replaceField.setText("NOMINA DEL MES DE: "+proximoMesSpanish.toUpperCase());

        JPanel inputPanel = new JPanel();
        inputPanel.setLayout(new FlowLayout());
        inputPanel.add(searchLabel);
        inputPanel.add(searchField);
        inputPanel.add(replaceLabel);
        inputPanel.add(replaceField);

        frame.add(inputPanel, BorderLayout.NORTH);

        errorLog.setWrapStyleWord(true);
        frame.add(new JScrollPane(errorLog), BorderLayout.CENTER);
        logger.addHandler(new TextAreaHandler(errorLog));
        errorLog.setCaretPosition(errorLog.getDocument().getLength());
        errorLog.revalidate();
        errorLog.repaint();
        frame.setVisible(true);

        JProgressBar progressBar = new JProgressBar();
        progressBar.setStringPainted(true);
        progressBar.setPreferredSize(new Dimension(850, 25));
        frame.add(progressBar, BorderLayout.SOUTH);

        JButton updateButton = new JButton("Choose directory");
        inputPanel.add(updateButton, BorderLayout.AFTER_LAST_LINE);

        updateButton.addActionListener(e -> {
            String userDir = System.getProperty("user.home");
            JFileChooser chooser = new JFileChooser(userDir +"/Desktop/CARPETA MES");
            chooser.setDialogTitle("Choose the folder containing Words documents for replacement");

            chooser.setFileSelectionMode(JFileChooser.DIRECTORIES_ONLY);
            int returnVal = chooser.showOpenDialog(null);
            if(returnVal == JFileChooser.APPROVE_OPTION) {
                folder = chooser.getSelectedFile();
                String searchPhrase = searchField.getText();
                String replacePhrase = replaceField.getText();
                // Iterate through the contents of the folder
                int dialogResult = JOptionPane.showConfirmDialog(frame, "We replace in all Words documents using "+replacePhrase, "Do you want to confirm the action?", dialogButton);
                if(dialogResult == 0) {
                    new Thread(() -> {
                        progressBar.setMaximum(FileCounter.countFiles(folder.toString(),".docx"));
                        progressBar.setValue(0);
                        logger.info("------------------------  Starting replacement in all documents...  ------------------------");
                        iterateFolder(folder,searchPhrase,replacePhrase,progressBar);
                        int dialogResult2 = JOptionPane.showConfirmDialog(frame, "Do you want to generate PDFS of the modified Words now??", "Generating PDFs", dialogButton);
                        if(dialogResult2 == 0) {
                            logger.info("------------------------  We attempt to create PDFS  ------------------------");
                            new Thread(() -> {
                                progressBar.setMaximum(FileCounter.countFiles(folder.toString(),".pdf"));
                                progressBar.setValue(0);
                                counter = 0;
                                CreatePDFS(folder,progressBar);
                                invokeLater(() -> {
                                    errorLog.revalidate();
                                    errorLog.repaint();
                                });
                                int dialogResult3 = JOptionPane.showConfirmDialog(frame, "¿Do you want to delete now ALL the docx templates that are no longer needed??", "Delete DOCXs templates", dialogButton);
                                if(dialogResult3 == 0) {
                                    logger.info("------------------------  Deleting docx templates that are no longer useful  ------------------------");
                                    new Thread(() -> {
                                        progressBar.setMaximum(FileCounter.countFiles(folder.toString(),".docx"));
                                        progressBar.setValue(0);
                                        counter = 0;
                                        invokeLater(() -> {
                                            errorLog.revalidate();
                                            errorLog.repaint();
                                            progressBar.revalidate();
                                            progressBar.repaint();

                                        });
                                        DeleteDocx(folder,progressBar);

                                    }).start();
                                }
                                JOptionPane.showMessageDialog(frame,
                                        "Process completed! You can now close the application.");
                                logger.info("------------------------  Process completed! You can now close the application. ------------------------  ");
                                invokeLater(() -> {
                                    errorLog.revalidate();
                                    errorLog.repaint();
                                });
                            }).start();
                        }
                        invokeLater(() -> {
                            errorLog.revalidate();
                            errorLog.repaint();
                        });
                    }).start();
                }
            }
        });
    }
    private static void iterateFolder(File folder,String searchPhrase,String replacePhrase,JProgressBar progressBar){

        File[] files = folder.listFiles();
        for (File file : Objects.requireNonNull(files)) {
            if (file.isDirectory()) {
                iterateFolder(file,searchPhrase,replacePhrase,progressBar);
            } else if (file.getName().endsWith(".docx")) {
                try {
                    // Load the docx file
                    WordprocessingMLPackage wordMLPackage = WordprocessingMLPackage.load(file);
                    // Build a list of "text" elements
                    List<Object> texts = wordMLPackage.getMainDocumentPart().getJAXBNodesViaXPath(XPATH_TO_SELECT_TEXT_NODES, true);

                    // Loop through all "text" elements
                    for (Object obj : texts) {
                        Text text;
                        text = (Text) ((JAXBElement<?>) obj).getValue();

                        // Get the text value
                        String textValueBefore = text.getValue();

                        // Perform the replacement
                        String textValueAfter = textValueBefore.replaceAll(searchPhrase, replacePhrase);
                        if(!textValueBefore.equals(textValueAfter)) {
                            logger.info("Modified Word: "+ file.getAbsoluteFile());
                            // Update the text element now that we have performed the replacement
                            text.setValue(textValueAfter);
                        }
                    }
                    wordMLPackage.save(file);
                } catch (Exception e) {
                    logger.severe(e.getMessage());
                    e.printStackTrace();
                } finally {
                    progressBar.setValue(++counter);
                }


            }
        }//end for
    }
    private static void CreatePDFS(File folder,JProgressBar progressBar) {
        try {
            walk(folder.toPath())
                    .filter(Files::isRegularFile)
                    .filter(path -> path.toString().endsWith(".docx"))
                    .forEach(path -> {
                        try {
                            InputStream docFile = Files.newInputStream(path.toFile().toPath());
                            OutputStream out;
                            try (XWPFDocument doc = new XWPFDocument(docFile)) {
                                PdfOptions pdfOptions = PdfOptions.create();
                                out = Files.newOutputStream(new File(path.toString().replace(".docx", "") + ".pdf").toPath());
                                PdfConverter.getInstance().convert(doc, out, pdfOptions);
                            }
                            out.close();
                            logger.info("PDF generated: " +path.toString().replace(".docx","")+".pdf");
                        }
                        catch(Exception e) {
                            e.printStackTrace();
                        }
                        finally {
                            progressBar.setValue(++counter);
                        }
                    });
        } catch (IOException e) {
            e.printStackTrace();
        }
    }
    private static void DeleteDocx(File folder,JProgressBar progressBar) {
        try {
            walk(folder.toPath())
                    .filter(Files::isRegularFile)
                    .filter(path -> path.toString().endsWith(".docx"))
                    .forEach(path -> {
                        try {
                            if (Files.exists(path)) {
                                try {
                                    // Delete the file
                                    Files.delete(path);
                                    logger.info("Deleted: " +path);
                                } catch (IOException e) {
                                    logger.info("Error while deleting: "+path);
                                    e.printStackTrace();
                                }
                            } else {
                                logger.info("File does not exist: " +path);
                            }
                        }
                        catch(Exception e) {
                            logger.severe(e.getMessage());
                            e.printStackTrace();
                        }finally {
                            progressBar.setValue(++counter);
                        }
                    });
        } catch (IOException e) {
            logger.severe(e.getMessage());
            e.printStackTrace();
        }
    }
    @Override
    public void actionPerformed(ActionEvent e) {
    }
}

