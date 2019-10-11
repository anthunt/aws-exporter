package anthunt.aws.exporter;

import org.eclipse.aether.AbstractRepositoryListener;
import org.eclipse.aether.RepositoryEvent;

public class ConsoleRepositoryEventListener extends AbstractRepositoryListener {

	@Override
	public void artifactInstalled(RepositoryEvent event) {
	    //System.out.printf("artifact %s installed to file %s\n, event.getArtifact()", event.getFile());
	}
	
	@Override
	public void artifactInstalling(RepositoryEvent event) {
	    //System.out.printf("installing artifact %s to file %s\n, event.getArtifact()", event.getFile());
	}
	
	@Override
	public void artifactResolved(RepositoryEvent event) {
	    //System.out.printf("artifact %s resolved from repository %s\n", event.getArtifact(), event.getRepository());
	}
	
	@Override
	public void artifactDownloading(RepositoryEvent event) {
	    //System.out.printf("downloading artifact %s from repository %s\n", event.getArtifact(), event.getRepository());
	}
	
	@Override
	public void artifactDownloaded(RepositoryEvent event) {
	    //System.out.printf("downloaded artifact %s from repository %s\n", event.getArtifact(), event.getRepository());
	}
	
	@Override
	public void artifactResolving(RepositoryEvent event) {
	    //System.out.printf("resolving artifact %s\n", event.getArtifact());
	}

}