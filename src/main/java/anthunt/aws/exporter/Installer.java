package anthunt.aws.exporter;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.net.URISyntaxException;

import org.apache.maven.model.building.ModelBuildingException;
import org.eclipse.aether.resolution.DependencyResolutionException;
import org.eclipse.aether.util.artifact.JavaScopes;

import anthunt.mvn.downloader.MvnDownloader;

public class Installer {

	public static void install() throws URISyntaxException, DependencyResolutionException, ModelBuildingException {
						
		MvnDownloader mvnDownloader = new MvnDownloader(getResourceAsFile("/pom.xml"), "lib", MvnDownloader.DEFAULT_MAVEN_REPOSITORY, JavaScopes.RUNTIME, new ConsoleRepositoryEventListener());
		
		mvnDownloader.getAllDependencies();
	}
	
	public static File getResourceAsFile(String resourcePath) {
	    try {
	        InputStream in = Installer.class.getResourceAsStream(resourcePath);
	        if (in == null) {
	            return null;
	        }

	        File tempFile = File.createTempFile(String.valueOf(in.hashCode()), ".tmp");
	        tempFile.deleteOnExit();

	        try (FileOutputStream out = new FileOutputStream(tempFile)) {
	            byte[] buffer = new byte[1024];
	            int bytesRead;
	            while ((bytesRead = in.read(buffer)) != -1) {
	                out.write(buffer, 0, bytesRead);
	            }
	        }
	        return tempFile;
	    } catch (IOException e) {
	        e.printStackTrace();
	        return null;
	    }
	}
	
	public static void main(String[] args) throws DependencyResolutionException, URISyntaxException, ModelBuildingException {
		Installer.install();
	}
}
