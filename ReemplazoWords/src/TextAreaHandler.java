import javax.swing.*;
import java.util.logging.Handler;
import java.util.logging.LogRecord;

import static javax.swing.SwingUtilities.invokeLater;

class TextAreaHandler extends Handler {
    private final JTextArea textArea;

    public TextAreaHandler(JTextArea textArea) {
        this.textArea = textArea;
        setFormatter(new TextAreaFormatter());
    }

    @Override
    public void publish(LogRecord record) {
        invokeLater(() -> {
            textArea.append(getFormatter().format(record));
            textArea.setCaretPosition(textArea.getDocument().getLength());
        });

    }
    @Override
    public void flush() { }

    @Override
    public void close() throws SecurityException { }
}
