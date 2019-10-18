package anthunt.aws.exporter;

import com.amazonaws.ClientConfiguration;
import com.amazonaws.auth.AWSCredentials;
import com.amazonaws.auth.AWSStaticCredentialsProvider;
import com.amazonaws.auth.BasicAWSCredentials;
import com.amazonaws.auth.BasicSessionCredentials;
import com.amazonaws.regions.Regions;
import com.amazonaws.services.apigateway.AmazonApiGateway;
import com.amazonaws.services.apigateway.AmazonApiGatewayClientBuilder;
import com.amazonaws.services.certificatemanager.AWSCertificateManager;
import com.amazonaws.services.certificatemanager.AWSCertificateManagerClientBuilder;
import com.amazonaws.services.directconnect.AmazonDirectConnect;
import com.amazonaws.services.directconnect.AmazonDirectConnectClientBuilder;
import com.amazonaws.services.directory.AWSDirectoryService;
import com.amazonaws.services.directory.AWSDirectoryServiceClientBuilder;
import com.amazonaws.services.ec2.AmazonEC2;
import com.amazonaws.services.ec2.AmazonEC2ClientBuilder;
import com.amazonaws.services.elasticache.AmazonElastiCache;
import com.amazonaws.services.elasticache.AmazonElastiCacheClientBuilder;
import com.amazonaws.services.elasticloadbalancing.AmazonElasticLoadBalancing;
import com.amazonaws.services.elasticloadbalancing.AmazonElasticLoadBalancingClientBuilder;
import com.amazonaws.services.identitymanagement.AmazonIdentityManagement;
import com.amazonaws.services.identitymanagement.AmazonIdentityManagementClientBuilder;
import com.amazonaws.services.kms.AWSKMS;
import com.amazonaws.services.kms.AWSKMSClientBuilder;
import com.amazonaws.services.lambda.AWSLambda;
import com.amazonaws.services.lambda.AWSLambdaClientBuilder;
import com.amazonaws.services.rds.AmazonRDS;
import com.amazonaws.services.rds.AmazonRDSClientBuilder;
import com.amazonaws.services.route53.AmazonRoute53;
import com.amazonaws.services.route53.AmazonRoute53ClientBuilder;
import com.amazonaws.services.s3.AmazonS3;
import com.amazonaws.services.s3.AmazonS3ClientBuilder;
import com.amazonaws.services.securitytoken.AWSSecurityTokenService;
import com.amazonaws.services.securitytoken.AWSSecurityTokenServiceClientBuilder;
import com.amazonaws.services.securitytoken.model.AssumeRoleRequest;
import com.amazonaws.services.securitytoken.model.AssumeRoleResult;
import com.amazonaws.services.securitytoken.model.Credentials;

import anthunt.aws.exporter.model.AmazonAccess;
import anthunt.aws.exporter.model.CrossAccountRole;

public class AmazonClients
{
	public AmazonEC2 AmazonEC2;
	public AmazonElasticLoadBalancing AmazonElasticLoadBalancing;
	public com.amazonaws.services.elasticloadbalancingv2.AmazonElasticLoadBalancing AmazonElasticLoadBalancing2;
	public AmazonElastiCache AmazonElastiCache;
	public AmazonRDS AmazonRDS;
	public AWSKMS AwsKMS;
	public AWSCertificateManager AwsCertificateManager;
	public AmazonS3 AmazonS3;
	public AWSLambda AwsLambda;
	public AmazonApiGateway AmazonApiGateway;
	public AmazonIdentityManagement AmazonIdentityManagement;
	public AmazonDirectConnect AmazonDirectConnect;
	public AWSDirectoryService AwsDirectoryService;
	public AmazonRoute53 AmazonRoute53;
    
	public AmazonClients(Regions regions, AmazonAccess amazonAccess, boolean isCrossAccountRole, String crossAccountKey) {
		ClientConfiguration config = new ClientConfiguration();
		if (amazonAccess.isUseProxy().booleanValue()) {
			config.setProxyHost(amazonAccess.getProxyHost());
			config.setProxyPort(amazonAccess.getProxyPort().intValue());
		}
    
		AWSCredentials credentials = new BasicAWSCredentials(amazonAccess.getAccessKey(), amazonAccess.getSecretKey());
		if(isCrossAccountRole) {
			CrossAccountRole crossAccountRole = amazonAccess.getCrossAccountRole(crossAccountKey);
			credentials = this.getAssumeRole(credentials, crossAccountRole.getCrossRoleArn(), crossAccountRole.getCrossRoleSessionName(), crossAccountRole.getExternId());
		}
    
		initial(config, regions, new AWSStaticCredentialsProvider(credentials));
    
	}
  
	private void initial(ClientConfiguration config, Regions regions, AWSStaticCredentialsProvider awsStaticCredentialsProvider) {
	    
	    this.AmazonEC2 = AmazonEC2ClientBuilder.standard()
	    										.withCredentials(awsStaticCredentialsProvider)
	    										.withClientConfiguration(config)
	    										.withRegion(regions)
	    										.build();
	    
	    this.AmazonElastiCache = AmazonElastiCacheClientBuilder.standard()
	    										.withCredentials(awsStaticCredentialsProvider)
	    										.withClientConfiguration(config)
	    										.withRegion(regions)
	    										.build();
	    
	    this.AmazonElasticLoadBalancing = AmazonElasticLoadBalancingClientBuilder.standard()
												.withCredentials(awsStaticCredentialsProvider)
												.withClientConfiguration(config)
												.withRegion(regions)
												.build();
	    
	    this.AmazonElasticLoadBalancing2 = com.amazonaws.services.elasticloadbalancingv2.AmazonElasticLoadBalancingClientBuilder.standard()
	    										.withCredentials(awsStaticCredentialsProvider)
	    										.withClientConfiguration(config)
	    										.withRegion(regions)
	    										.build();
	    
	    this.AmazonRDS = AmazonRDSClientBuilder.standard()
	    										.withClientConfiguration(config)
	    										.withCredentials(awsStaticCredentialsProvider)
	    										.withRegion(regions)
	    										.build();
	    
	    this.AwsKMS = AWSKMSClientBuilder.standard()
												.withClientConfiguration(config)
												.withCredentials(awsStaticCredentialsProvider)
												.withRegion(regions)
												.build();
	    
	    this.AwsCertificateManager = AWSCertificateManagerClientBuilder.standard()
												.withClientConfiguration(config)
												.withCredentials(awsStaticCredentialsProvider)
												.withRegion(regions)
												.build();
	    
	    this.AmazonS3 = AmazonS3ClientBuilder.standard()
												.withClientConfiguration(config)
												.withCredentials(awsStaticCredentialsProvider)
												.withRegion(regions)
												.build();
	    
	    this.AwsLambda = AWSLambdaClientBuilder.standard()
												.withClientConfiguration(config)
												.withCredentials(awsStaticCredentialsProvider)
												.withRegion(regions)
												.build();
	    
	    this.AmazonApiGateway = AmazonApiGatewayClientBuilder.standard()
												.withClientConfiguration(config)
												.withCredentials(awsStaticCredentialsProvider)
												.withRegion(regions)
												.build();
	    
	    this.AmazonIdentityManagement = AmazonIdentityManagementClientBuilder.standard()
												.withClientConfiguration(config)
												.withCredentials(awsStaticCredentialsProvider)
												.withRegion(regions)
												.build();
	    
	    this.AmazonDirectConnect = AmazonDirectConnectClientBuilder.standard()
												.withClientConfiguration(config)
												.withCredentials(awsStaticCredentialsProvider)
												.withRegion(regions)
												.build();
	    
	    this.AwsDirectoryService = AWSDirectoryServiceClientBuilder.standard()
												.withClientConfiguration(config)
												.withCredentials(awsStaticCredentialsProvider)
												.withRegion(regions)
												.build();
	    
	    this.AmazonRoute53 = AmazonRoute53ClientBuilder.standard()
												.withClientConfiguration(config)
												.withCredentials(awsStaticCredentialsProvider)
												.withRegion(regions)
												.build();
	    
	}
  
	private AWSCredentials getAssumeRole(AWSCredentials credentials, String crossAccountRoleArn, String crossAccountRoleSessionName, String crossAccountRoleExternalId) {
	  
		AWSSecurityTokenService awsSecurityTokenService = AWSSecurityTokenServiceClientBuilder.standard()
															.withCredentials(new AWSStaticCredentialsProvider(credentials))
														    .withRegion(Regions.DEFAULT_REGION)
														    .build();
		
		AssumeRoleRequest assumeRoleRequest = new AssumeRoleRequest()
				.withRoleArn(crossAccountRoleArn)
				.withRoleSessionName(crossAccountRoleSessionName)
				.withExternalId(crossAccountRoleExternalId);
		
		AssumeRoleResult assumeRoleResult = awsSecurityTokenService.assumeRole(assumeRoleRequest);
				
		Credentials assumeRoleCredentials = assumeRoleResult.getCredentials();

		BasicSessionCredentials crossAccountSessionCredentials = new BasicSessionCredentials(
			  assumeRoleCredentials.getAccessKeyId()
			, assumeRoleCredentials.getSecretAccessKey()
			, assumeRoleCredentials.getSessionToken()
		);
		
		return crossAccountSessionCredentials;
		
	}
  
}
