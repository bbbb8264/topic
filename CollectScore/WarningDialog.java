package sd;
import javax.swing.JFrame;
import javax.swing.JOptionPane;

/***************************
 * ?�m?�X�����ءA?�Τ_ĵ�i
 * ************************/
public class WarningDialog extends JFrame{
    public WarningDialog(String str){
        JOptionPane.showMessageDialog(this,str, "Warning",JOptionPane.WARNING_MESSAGE);
    }
}