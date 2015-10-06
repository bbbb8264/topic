package sd;
import java.io.File;
import javax.swing.filechooser.FileFilter;

public class MyFileFilter extends FileFilter {
	String extension;
	String description;
	String extension2;

	public MyFileFilter(String ext,String ext2, String desc){
		this.extension = ext;
		extension2 = ext2;
		this.description = desc;
	}
	public MyFileFilter(String desc){
		this.description = desc;
	}
	public static String getExtension(File f) {
		String ext = null;
		String s = f.getName();
		int i = s.lastIndexOf('.');
		
		if (i > 0 && i < s.length() - 1) {
			ext = s.substring(i+1).toLowerCase();
		}
		return ext;
	}

	public boolean accept(File f) {
	if (f.isDirectory()) {
		return true;
	}
	String ext = this.getExtension(f);
	if (ext != null) {
		if (ext.equals(this.extension) || ext.equals(extension2)) 
		{
			return true;
		} else {
			return false;
		}
	}
	return false;
	}
	
	public String getDescription() {
		return this.description;
	}
}
