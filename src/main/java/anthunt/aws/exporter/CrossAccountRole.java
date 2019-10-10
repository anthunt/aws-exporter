package anthunt.aws.exporter;

public class CrossAccountRole {
	
	private String crossAccountId;
	private String crossRoleName;
	private String crossRoleSessionName;
	private String externId;
	
	public CrossAccountRole() {}
	
	public CrossAccountRole(String crossAccountId, String crossRoleName) {
		this(crossAccountId, crossRoleName, crossAccountId + "@" + crossRoleName);
	}

	public CrossAccountRole(String crossAccountId, String crossRoleName, String crossRoleSessionName) {
		this.setCrossAccountId(crossAccountId);
		this.setCrossRoleName(crossRoleName);
		this.setCrossRoleSessionName(crossRoleSessionName);
	}
	
	public String getCrossAccountId() {
		return crossAccountId;
	}
	
	public void setCrossAccountId(String crossAccountId) {
		this.crossAccountId = crossAccountId;
	}
	
	public String getCrossRoleName() {
		return crossRoleName;
	}
	
	public void setCrossRoleName(String crossRoleName) {
		this.crossRoleName = crossRoleName;
	}
	
	public String getCrossRoleSessionName() {
		return crossRoleSessionName;
	}
	
	public void setCrossRoleSessionName(String crossRoleSessionName) {
		this.crossRoleSessionName = crossRoleSessionName;
	}
	
	public String getExternId() {
		return "".equals(externId) ? null : externId;
	}
	
	public void setExternId(String externId) {
		this.externId = externId;
	}
	
	public String getCrossRoleArn() {
		StringBuilder crossRoleArn = new StringBuilder();
		
		crossRoleArn.append("arn:aws:iam::")
					.append(this.getCrossAccountId())
					.append(":role/")
					.append(this.getCrossRoleName());
		
		return crossRoleArn.toString();
	}

}
