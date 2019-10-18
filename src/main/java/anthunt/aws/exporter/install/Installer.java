package anthunt.aws.exporter.install;

import java.io.File;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.net.URISyntaxException;

import org.apache.maven.model.building.ModelBuildingException;
import org.eclipse.aether.resolution.DependencyResolutionException;
import org.eclipse.aether.util.artifact.JavaScopes;

import anthunt.mvn.downloader.MvnDownloader;

public class Installer {

	private static final String LIB_PATH = "lib";
	private static final String POM_PATH = "/pom.xml";
	private static final String SCRIPT_PATH = "/bin/startAWSExporter.bat";
	
	private static void install() throws URISyntaxException, DependencyResolutionException, ModelBuildingException {
		
		File pomFile = getResourceAsTempFile(POM_PATH);
		File libDir = new File(LIB_PATH);
		
		System.out.println("Start Dependancies to lib directory");
		System.out.println("library jar install to " + libDir.getAbsolutePath());
		
		MvnDownloader mvnDownloader = new MvnDownloader(pomFile, LIB_PATH, MvnDownloader.DEFAULT_MAVEN_REPOSITORY, JavaScopes.PROVIDED, new InstallEventListener());
		mvnDownloader.getAllDependencies();
		System.out.println("resolved all dependancies");
		System.out.println("start library configuration");
		moveJars(libDir);		
		makeRunScript();		
		System.out.println("end of install");
		
	}
	
	private static void moveJars(File dir) {
		for(String path : dir.list()) {
			File child = new File(dir, path);
			if(child.isDirectory()) {
				moveJars(child);
			} else {
				if(child.getName().endsWith(".jar")) {
					System.out.printf("resolved library : %s, %s\n", child.getName(), child.renameTo(new File(LIB_PATH, child.getName())));
				}
			}
			child.delete();
		}
	}
	
	private static void makeRunScript() {
		getResourceAsFile(SCRIPT_PATH);
	}
	
	private static File getResourceAsTempFile(String resourcePath) {
	    try {
	        InputStream in = Installer.class.getResourceAsStream(resourcePath);
	        if (in == null) {
	            return null;
	        }

	        File tempFile = File.createTempFile(String.valueOf(in.hashCode()), ".tmp");
	        tempFile.deleteOnExit();
	        makeFile(in, tempFile);
	        return tempFile;
	    } catch (IOException e) {
	        e.printStackTrace();
	        return null;
	    }
	}

	private static File getResourceAsFile(String resourcePath) {
	    try {
	        InputStream in = Installer.class.getResourceAsStream(resourcePath);
	        if (in == null) {
	            return null;
	        }

	        File runFile = new File(resourcePath.substring(resourcePath.lastIndexOf("/") + 1, resourcePath.length()));
	        makeFile(in, runFile);
	        return runFile;
	    } catch (IOException e) {
	        e.printStackTrace();
	        return null;
	    }
	}
	
	private static void makeFile(InputStream in, File file) throws FileNotFoundException, IOException {
		try (FileOutputStream out = new FileOutputStream(file)) {
            byte[] buffer = new byte[1024];
            int bytesRead;
            while ((bytesRead = in.read(buffer)) != -1) {
                out.write(buffer, 0, bytesRead);
            }
        }
	}
	
	public static void main(String[] args) throws DependencyResolutionException, URISyntaxException, ModelBuildingException {
		install();
	}
}
