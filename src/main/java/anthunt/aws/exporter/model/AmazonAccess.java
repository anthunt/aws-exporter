package anthunt.aws.exporter.model;

import java.util.HashMap;
import java.util.Map;

public class AmazonAccess
{
private String accessType;
private String accessKey;
private String secretKey;
private Boolean useProxy;
private String proxyHost;
private Integer proxyPort;
private Map<String, CrossAccountRole> crossAccountRoles;
  
	public AmazonAccess(String accessType)
	{
	    this.accessType = accessType;
	    this.crossAccountRoles = new HashMap<>();
	}
	  
	public String getAccessType()
	{
	    return this.accessType;
	}
	  
	public String getAccessKey()
	{
	    return this.accessKey;
	}
	  
	public void setAccessKey(String accessKey)
	{
	    this.accessKey = accessKey;
	}
	  
	public String getSecretKey()
	{
	    return this.secretKey;
	}
	  
	public void setSecretKey(String secretKey)
	{
	    this.secretKey = secretKey;
	}
	  
	public Boolean isUseProxy()
	{
	    return this.useProxy;
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

	public CrossAccountRole getCrossAccountRole(String crossAccountRoleName) {
		
		if(!this.crossAccountRoles.containsKey(crossAccountRoleName)) {
			this.crossAccountRoles.put(crossAccountRoleName, new CrossAccountRole());
		}
		return this.crossAccountRoles.get(crossAccountRoleName);
	}
	
	public Map<String, CrossAccountRole> getCrossAccountRoles() {
		return crossAccountRoles;
	}
	
	public void setCrossAccountRoles(Map<String, CrossAccountRole> crossAccountRoles) {
		this.crossAccountRoles = crossAccountRoles;
	}
	
}
