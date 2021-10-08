package anthunt.aws.exporter;

import anthunt.aws.exporter.model.AmazonAccess;
import org.apache.http.impl.client.BasicCredentialsProvider;
import software.amazon.awssdk.auth.credentials.*;
import software.amazon.awssdk.profiles.Profile;
import software.amazon.awssdk.profiles.ProfileFile;
import software.amazon.awssdk.regions.Region;
import software.amazon.awssdk.services.acm.AcmClient;
import software.amazon.awssdk.services.apigateway.ApiGatewayClient;
import software.amazon.awssdk.services.directconnect.DirectConnectClient;
import software.amazon.awssdk.services.directory.DirectoryClient;
import software.amazon.awssdk.services.ec2.Ec2Client;
import software.amazon.awssdk.services.elasticache.ElastiCacheClient;
import software.amazon.awssdk.services.elasticloadbalancing.ElasticLoadBalancingClient;
import software.amazon.awssdk.services.elasticloadbalancingv2.ElasticLoadBalancingV2Client;
import software.amazon.awssdk.services.iam.IamClient;
import software.amazon.awssdk.services.kms.KmsClient;
import software.amazon.awssdk.services.lambda.LambdaClient;
import software.amazon.awssdk.services.rds.RdsClient;
import software.amazon.awssdk.services.route53.Route53Client;
import software.amazon.awssdk.services.s3.S3Client;
import software.amazon.awssdk.services.sts.StsClient;
import software.amazon.awssdk.services.sts.model.AssumeRoleRequest;
import software.amazon.awssdk.services.sts.model.Credentials;

import java.util.Optional;
import java.util.Scanner;
import java.util.function.Supplier;

public class AmazonClients
{
	public Ec2Client ec2Client;
	public ElasticLoadBalancingClient elasticLoadBalancingClient;
	public ElasticLoadBalancingV2Client elasticLoadBalancingV2Client;
	public ElastiCacheClient elastiCacheClient;
	public RdsClient rdsClient;
	public KmsClient kmsClient;
	public AcmClient acmClient;
	public S3Client s3Client;
	public LambdaClient lambdaClient;
	public ApiGatewayClient apiGatewayClient;
	public IamClient iamClient;
	public DirectConnectClient directConnectClient;
	public DirectoryClient directoryClient;
	public Route53Client route53Client;

	public AmazonClients(AmazonAccess amazonAccess, String profileName, Region region) {

		/*
		ClientConfiguration config = new ClientConfiguration();
		if (amazonAccess.isUseProxy().booleanValue()) {
			config.setProxyHost(amazonAccess.getProxyHost());
			config.setProxyPort(amazonAccess.getProxyPort().intValue());
		}
    	*/

		Supplier<ProfileFile> defaultProfileFileLoader = ProfileFile::defaultProfileFile;
		ProfileFile profileFile = defaultProfileFileLoader.get();

		Profile profile = profileFile.profile(profileName).get();
		String source_profile = profile.property("source_profile").get();

		Profile sourceProfile = profileFile.profile(source_profile).get();

		AwsBasicCredentials awsBasicCredentials = AwsBasicCredentials.create(
				sourceProfile.property("aws_access_key_id").get(),
				sourceProfile.property("aws_secret_access_key").get()
		);

		StsClient stsClient = StsClient.builder()
				.region(Region.of(sourceProfile.property("region").get()))
				.credentialsProvider(StaticCredentialsProvider.create(awsBasicCredentials))
				.build();

		Scanner sc = null;
		String tokenCode = "";
		try {
			sc = new Scanner(System.in);
			System.out.print("MFA code : ");
			while (!sc.hasNextLine()) {
				sc.next();
			}
			tokenCode = sc.nextLine();
		} catch (Exception skip) {}
		System.out.println("TokenCode = " + tokenCode);

		Credentials credentials = stsClient.assumeRole(
				AssumeRoleRequest.builder()
						.roleArn(profile.property("role_arn").get())
						.roleSessionName(profileName)
						.serialNumber(profile.property("mfa_serial").get())
						.tokenCode(tokenCode)
						.build()
		).credentials();

		AwsSessionCredentials awsSessionCredentials = AwsSessionCredentials.create(
				credentials.accessKeyId(),
				credentials.secretAccessKey(),
				credentials.sessionToken()
		);

		initial(region, StaticCredentialsProvider.create(awsSessionCredentials)); //ProfileCredentialsProvider.create(profileName));

	}

	private void initial(Region region, AwsCredentialsProvider profileCredentialsProvider) {

		this.ec2Client = Ec2Client.builder()
				.region(region)
				.credentialsProvider(profileCredentialsProvider)
				.build();

		this.elastiCacheClient = ElastiCacheClient.builder()
				.region(region)
				.credentialsProvider(profileCredentialsProvider)
				.build();

		this.elasticLoadBalancingClient = ElasticLoadBalancingClient.builder()
				.region(region)
				.credentialsProvider(profileCredentialsProvider)
				.build();

		this.elasticLoadBalancingV2Client = ElasticLoadBalancingV2Client.builder()
				.region(region)
				.credentialsProvider(profileCredentialsProvider)
				.build();

		this.rdsClient = RdsClient.builder()
				.region(region)
				.credentialsProvider(profileCredentialsProvider)
				.build();

		this.kmsClient = KmsClient.builder()
				.region(region)
				.credentialsProvider(profileCredentialsProvider)
				.build();

		this.acmClient = AcmClient.builder()
				.region(region)
				.credentialsProvider(profileCredentialsProvider)
				.build();

		this.s3Client = S3Client.builder()
				.region(region)
				.credentialsProvider(profileCredentialsProvider)
				.build();

		this.lambdaClient = LambdaClient.builder()
				.region(region)
				.credentialsProvider(profileCredentialsProvider)
				.build();

		this.apiGatewayClient = ApiGatewayClient.builder()
				.region(region)
				.credentialsProvider(profileCredentialsProvider)
				.build();

		this.iamClient = IamClient.builder()
				.region(region)
				.credentialsProvider(profileCredentialsProvider)
				.build();

		this.directConnectClient = DirectConnectClient.builder()
				.region(region)
				.credentialsProvider(profileCredentialsProvider)
				.build();

		this.directoryClient = DirectoryClient.builder()
				.region(region)
				.credentialsProvider(profileCredentialsProvider)
				.build();

		this.route53Client = Route53Client.builder()
				.region(region)
				.credentialsProvider(profileCredentialsProvider)
				.build();

	}

}
