package sd;
import javax.swing.JFrame;
import javax.swing.JOptionPane;

/***************************
 * ?置?出消息框，?用于警告
 * ************************/
public class WarningDialog extends JFrame{
    public WarningDialog(String str){
        JOptionPane.showMessageDialog(this,str, "Warning",JOptionPane.WARNING_MESSAGE);
    }
}