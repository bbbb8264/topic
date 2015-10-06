package v1;

import java.io.File;
import javax.swing.filechooser.FileFilter;

public class MyFileFilter extends FileFilter {
	String extension;
	String description;

	public MyFileFilter(String ext, String desc){
		this.extension = ext;
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
		if (ext.equals(this.extension)) 
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
