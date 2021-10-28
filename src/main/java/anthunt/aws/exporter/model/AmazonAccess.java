package anthunt.aws.exporter.model;

public class AmazonAccess
{
	private Boolean useProxy = false;
	private String proxyHost;
	private Integer proxyPort;
  
	public AmazonAccess() {}
	  	  
	public Boolean isUseProxy()
	{
	    return this.useProxy == null ? false : this.useProxy;
	}
	  
	public void setUseProxy(Boolean useProxy)
	{
	    this.useProxy = useProxy;
	}
	  
	public String getProxyHost()
	{
	    return this.proxyHost;
	}
	  
	public void setProxyHost(String proxyHost)
	{
	    this.proxyHost = proxyHost;
	}
	  
	public Integer getProxyPort()
	{
	    return this.proxyPort;
	}
	  
	public void setProxyPort(Integer proxyPort)
	{
	    this.proxyPort = proxyPort;
	}
	
}
