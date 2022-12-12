package anthunt.aws.exporter;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.text.DateFormat;
import java.text.SimpleDateFormat;
import java.time.ZoneOffset;
import java.time.format.DateTimeFormatter;
import java.util.ArrayList;
import java.util.Date;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import anthunt.aws.exporter.model.AmazonAccess;
import anthunt.poi.helper.XSSFHelper;
import software.amazon.awssdk.regions.Region;
import software.amazon.awssdk.services.acm.model.CertificateDetail;
import software.amazon.awssdk.services.acm.model.CertificateSummary;
import software.amazon.awssdk.services.acm.model.DescribeCertificateRequest;
import software.amazon.awssdk.services.acm.model.DescribeCertificateResponse;
import software.amazon.awssdk.services.acm.model.DomainValidation;
import software.amazon.awssdk.services.acm.model.ExtendedKeyUsage;
import software.amazon.awssdk.services.acm.model.KeyUsage;
import software.amazon.awssdk.services.acm.model.ListCertificatesResponse;
import software.amazon.awssdk.services.acm.model.RenewalSummary;
import software.amazon.awssdk.services.autoscaling.model.AutoScalingGroup;
import software.amazon.awssdk.services.autoscaling.model.DescribeAutoScalingGroupsResponse;
import software.amazon.awssdk.services.directconnect.model.BGPPeer;
import software.amazon.awssdk.services.directconnect.model.Connection;
import software.amazon.awssdk.services.directconnect.model.DescribeConnectionsResponse;
import software.amazon.awssdk.services.directconnect.model.DescribeDirectConnectGatewaysRequest;
import software.amazon.awssdk.services.directconnect.model.DescribeDirectConnectGatewaysResponse;
import software.amazon.awssdk.services.directconnect.model.DescribeLagsRequest;
import software.amazon.awssdk.services.directconnect.model.DescribeLagsResponse;
import software.amazon.awssdk.services.directconnect.model.DescribeLocationsResponse;
import software.amazon.awssdk.services.directconnect.model.DescribeVirtualGatewaysResponse;
import software.amazon.awssdk.services.directconnect.model.DescribeVirtualInterfacesResponse;
import software.amazon.awssdk.services.directconnect.model.DirectConnectGateway;
import software.amazon.awssdk.services.directconnect.model.Lag;
import software.amazon.awssdk.services.directconnect.model.Location;
import software.amazon.awssdk.services.directconnect.model.RouteFilterPrefix;
import software.amazon.awssdk.services.directconnect.model.VirtualGateway;
import software.amazon.awssdk.services.directconnect.model.VirtualInterface;
import software.amazon.awssdk.services.directory.model.ConditionalForwarder;
import software.amazon.awssdk.services.directory.model.DescribeConditionalForwardersRequest;
import software.amazon.awssdk.services.directory.model.DescribeConditionalForwardersResponse;
import software.amazon.awssdk.services.directory.model.DescribeDirectoriesResponse;
import software.amazon.awssdk.services.directory.model.DescribeDomainControllersRequest;
import software.amazon.awssdk.services.directory.model.DescribeDomainControllersResponse;
import software.amazon.awssdk.services.directory.model.DescribeTrustsRequest;
import software.amazon.awssdk.services.directory.model.DescribeTrustsResponse;
import software.amazon.awssdk.services.directory.model.DirectoryConnectSettingsDescription;
import software.amazon.awssdk.services.directory.model.DirectoryDescription;
import software.amazon.awssdk.services.directory.model.DirectoryVpcSettingsDescription;
import software.amazon.awssdk.services.directory.model.DomainController;
import software.amazon.awssdk.services.directory.model.InvalidParameterException;
import software.amazon.awssdk.services.directory.model.RadiusSettings;
import software.amazon.awssdk.services.directory.model.Trust;
import software.amazon.awssdk.services.ec2.model.CidrBlock;
import software.amazon.awssdk.services.ec2.model.CpuOptions;
import software.amazon.awssdk.services.ec2.model.CustomerGateway;
import software.amazon.awssdk.services.ec2.model.DescribeCustomerGatewaysResponse;
import software.amazon.awssdk.services.ec2.model.DescribeEgressOnlyInternetGatewaysRequest;
import software.amazon.awssdk.services.ec2.model.DescribeEgressOnlyInternetGatewaysResponse;
import software.amazon.awssdk.services.ec2.model.DescribeInstancesRequest;
import software.amazon.awssdk.services.ec2.model.DescribeNatGatewaysRequest;
import software.amazon.awssdk.services.ec2.model.DescribeNatGatewaysResponse;
import software.amazon.awssdk.services.ec2.model.DescribeNetworkAclsRequest;
import software.amazon.awssdk.services.ec2.model.DescribeNetworkAclsResponse;
import software.amazon.awssdk.services.ec2.model.DescribeRouteTablesRequest;
import software.amazon.awssdk.services.ec2.model.DescribeSecurityGroupsRequest;
import software.amazon.awssdk.services.ec2.model.DescribeSecurityGroupsResponse;
import software.amazon.awssdk.services.ec2.model.DescribeSubnetsRequest;
import software.amazon.awssdk.services.ec2.model.DescribeVpcPeeringConnectionsResponse;
import software.amazon.awssdk.services.ec2.model.DescribeVpnConnectionsResponse;
import software.amazon.awssdk.services.ec2.model.DescribeVpnGatewaysResponse;
import software.amazon.awssdk.services.ec2.model.EbsInstanceBlockDevice;
import software.amazon.awssdk.services.ec2.model.EgressOnlyInternetGateway;
import software.amazon.awssdk.services.ec2.model.ElasticGpuAssociation;
import software.amazon.awssdk.services.ec2.model.Filter;
import software.amazon.awssdk.services.ec2.model.GroupIdentifier;
import software.amazon.awssdk.services.ec2.model.IamInstanceProfile;
import software.amazon.awssdk.services.ec2.model.Instance;
import software.amazon.awssdk.services.ec2.model.InstanceBlockDeviceMapping;
import software.amazon.awssdk.services.ec2.model.InstanceIpv6Address;
import software.amazon.awssdk.services.ec2.model.InstanceNetworkInterface;
import software.amazon.awssdk.services.ec2.model.InstanceNetworkInterfaceAssociation;
import software.amazon.awssdk.services.ec2.model.InstanceNetworkInterfaceAttachment;
import software.amazon.awssdk.services.ec2.model.InstancePrivateIpAddress;
import software.amazon.awssdk.services.ec2.model.InstanceState;
import software.amazon.awssdk.services.ec2.model.InternetGateway;
import software.amazon.awssdk.services.ec2.model.InternetGatewayAttachment;
import software.amazon.awssdk.services.ec2.model.IpPermission;
import software.amazon.awssdk.services.ec2.model.IpRange;
import software.amazon.awssdk.services.ec2.model.Ipv6CidrBlock;
import software.amazon.awssdk.services.ec2.model.Ipv6Range;
import software.amazon.awssdk.services.ec2.model.Monitoring;
import software.amazon.awssdk.services.ec2.model.NatGateway;
import software.amazon.awssdk.services.ec2.model.NatGatewayAddress;
import software.amazon.awssdk.services.ec2.model.NetworkAcl;
import software.amazon.awssdk.services.ec2.model.Placement;
import software.amazon.awssdk.services.ec2.model.PrefixListId;
import software.amazon.awssdk.services.ec2.model.ProductCode;
import software.amazon.awssdk.services.ec2.model.ProvisionedBandwidth;
import software.amazon.awssdk.services.ec2.model.Reservation;
import software.amazon.awssdk.services.ec2.model.Route;
import software.amazon.awssdk.services.ec2.model.RouteTable;
import software.amazon.awssdk.services.ec2.model.RouteTableAssociation;
import software.amazon.awssdk.services.ec2.model.SecurityGroup;
import software.amazon.awssdk.services.ec2.model.StateReason;
import software.amazon.awssdk.services.ec2.model.Subnet;
import software.amazon.awssdk.services.ec2.model.SubnetCidrBlockState;
import software.amazon.awssdk.services.ec2.model.SubnetIpv6CidrBlockAssociation;
import software.amazon.awssdk.services.ec2.model.Tag;
import software.amazon.awssdk.services.ec2.model.UserIdGroupPair;
import software.amazon.awssdk.services.ec2.model.VgwTelemetry;
import software.amazon.awssdk.services.ec2.model.Volume;
import software.amazon.awssdk.services.ec2.model.VolumeAttachment;
import software.amazon.awssdk.services.ec2.model.Vpc;
import software.amazon.awssdk.services.ec2.model.VpcAttachment;
import software.amazon.awssdk.services.ec2.model.VpcCidrBlockAssociation;
import software.amazon.awssdk.services.ec2.model.VpcCidrBlockState;
import software.amazon.awssdk.services.ec2.model.VpcIpv6CidrBlockAssociation;
import software.amazon.awssdk.services.ec2.model.VpcPeeringConnection;
import software.amazon.awssdk.services.ec2.model.VpcPeeringConnectionOptionsDescription;
import software.amazon.awssdk.services.ec2.model.VpcPeeringConnectionStateReason;
import software.amazon.awssdk.services.ec2.model.VpcPeeringConnectionVpcInfo;
import software.amazon.awssdk.services.ec2.model.VpnConnection;
import software.amazon.awssdk.services.ec2.model.VpnGateway;
import software.amazon.awssdk.services.ec2.model.VpnStaticRoute;
import software.amazon.awssdk.services.elasticache.model.CacheCluster;
import software.amazon.awssdk.services.elasticache.model.DescribeCacheClustersRequest;
import software.amazon.awssdk.services.elasticache.model.NodeGroup;
import software.amazon.awssdk.services.elasticache.model.NodeGroupMember;
import software.amazon.awssdk.services.elasticache.model.ReplicationGroup;
import software.amazon.awssdk.services.elasticache.model.SecurityGroupMembership;
import software.amazon.awssdk.services.elasticloadbalancing.model.Listener;
import software.amazon.awssdk.services.elasticloadbalancing.model.ListenerDescription;
import software.amazon.awssdk.services.elasticloadbalancing.model.LoadBalancerDescription;
import software.amazon.awssdk.services.elasticloadbalancingv2.model.Action;
import software.amazon.awssdk.services.elasticloadbalancingv2.model.AvailabilityZone;
import software.amazon.awssdk.services.elasticloadbalancingv2.model.Certificate;
import software.amazon.awssdk.services.elasticloadbalancingv2.model.DescribeListenersRequest;
import software.amazon.awssdk.services.elasticloadbalancingv2.model.DescribeListenersResponse;
import software.amazon.awssdk.services.elasticloadbalancingv2.model.DescribeLoadBalancersResponse;
import software.amazon.awssdk.services.elasticloadbalancingv2.model.DescribeRulesRequest;
import software.amazon.awssdk.services.elasticloadbalancingv2.model.DescribeRulesResponse;
import software.amazon.awssdk.services.elasticloadbalancingv2.model.DescribeTagsRequest;
import software.amazon.awssdk.services.elasticloadbalancingv2.model.DescribeTagsResponse;
import software.amazon.awssdk.services.elasticloadbalancingv2.model.DescribeTargetGroupAttributesRequest;
import software.amazon.awssdk.services.elasticloadbalancingv2.model.DescribeTargetGroupAttributesResponse;
import software.amazon.awssdk.services.elasticloadbalancingv2.model.DescribeTargetGroupsRequest;
import software.amazon.awssdk.services.elasticloadbalancingv2.model.DescribeTargetGroupsResponse;
import software.amazon.awssdk.services.elasticloadbalancingv2.model.DescribeTargetHealthRequest;
import software.amazon.awssdk.services.elasticloadbalancingv2.model.DescribeTargetHealthResponse;
import software.amazon.awssdk.services.elasticloadbalancingv2.model.LoadBalancer;
import software.amazon.awssdk.services.elasticloadbalancingv2.model.LoadBalancerAddress;
import software.amazon.awssdk.services.elasticloadbalancingv2.model.Rule;
import software.amazon.awssdk.services.elasticloadbalancingv2.model.RuleCondition;
import software.amazon.awssdk.services.elasticloadbalancingv2.model.TagDescription;
import software.amazon.awssdk.services.elasticloadbalancingv2.model.TargetDescription;
import software.amazon.awssdk.services.elasticloadbalancingv2.model.TargetGroup;
import software.amazon.awssdk.services.elasticloadbalancingv2.model.TargetGroupAttribute;
import software.amazon.awssdk.services.elasticloadbalancingv2.model.TargetHealth;
import software.amazon.awssdk.services.elasticloadbalancingv2.model.TargetHealthDescription;
import software.amazon.awssdk.services.kms.model.AliasListEntry;
import software.amazon.awssdk.services.kms.model.DescribeKeyRequest;
import software.amazon.awssdk.services.kms.model.DescribeKeyResponse;
import software.amazon.awssdk.services.kms.model.KeyListEntry;
import software.amazon.awssdk.services.kms.model.KeyMetadata;
import software.amazon.awssdk.services.kms.model.ListAliasesResponse;
import software.amazon.awssdk.services.kms.model.ListKeysResponse;
import software.amazon.awssdk.services.rds.model.DBCluster;
import software.amazon.awssdk.services.rds.model.DBClusterMember;
import software.amazon.awssdk.services.rds.model.DBClusterOptionGroupStatus;
import software.amazon.awssdk.services.rds.model.DBClusterRole;
import software.amazon.awssdk.services.rds.model.DBInstance;
import software.amazon.awssdk.services.rds.model.DBParameterGroupStatus;
import software.amazon.awssdk.services.rds.model.DBSecurityGroupMembership;
import software.amazon.awssdk.services.rds.model.DescribeDbClustersResponse;
import software.amazon.awssdk.services.rds.model.DescribeDbInstancesResponse;
import software.amazon.awssdk.services.rds.model.DomainMembership;
import software.amazon.awssdk.services.rds.model.Endpoint;
import software.amazon.awssdk.services.rds.model.VpcSecurityGroupMembership;
import software.amazon.awssdk.services.s3.model.AbortIncompleteMultipartUpload;
import software.amazon.awssdk.services.s3.model.Bucket;
import software.amazon.awssdk.services.s3.model.CORSRule;
import software.amazon.awssdk.services.s3.model.Condition;
import software.amazon.awssdk.services.s3.model.Destination;
import software.amazon.awssdk.services.s3.model.GetBucketAccelerateConfigurationRequest;
import software.amazon.awssdk.services.s3.model.GetBucketAccelerateConfigurationResponse;
import software.amazon.awssdk.services.s3.model.GetBucketAclRequest;
import software.amazon.awssdk.services.s3.model.GetBucketAclResponse;
import software.amazon.awssdk.services.s3.model.GetBucketCorsRequest;
import software.amazon.awssdk.services.s3.model.GetBucketCorsResponse;
import software.amazon.awssdk.services.s3.model.GetBucketEncryptionRequest;
import software.amazon.awssdk.services.s3.model.GetBucketEncryptionResponse;
import software.amazon.awssdk.services.s3.model.GetBucketLifecycleConfigurationRequest;
import software.amazon.awssdk.services.s3.model.GetBucketLifecycleConfigurationResponse;
import software.amazon.awssdk.services.s3.model.GetBucketLocationRequest;
import software.amazon.awssdk.services.s3.model.GetBucketLoggingRequest;
import software.amazon.awssdk.services.s3.model.GetBucketLoggingResponse;
import software.amazon.awssdk.services.s3.model.GetBucketNotificationConfigurationRequest;
import software.amazon.awssdk.services.s3.model.GetBucketNotificationConfigurationResponse;
import software.amazon.awssdk.services.s3.model.GetBucketPolicyRequest;
import software.amazon.awssdk.services.s3.model.GetBucketReplicationRequest;
import software.amazon.awssdk.services.s3.model.GetBucketReplicationResponse;
import software.amazon.awssdk.services.s3.model.GetBucketTaggingRequest;
import software.amazon.awssdk.services.s3.model.GetBucketTaggingResponse;
import software.amazon.awssdk.services.s3.model.GetBucketVersioningRequest;
import software.amazon.awssdk.services.s3.model.GetBucketVersioningResponse;
import software.amazon.awssdk.services.s3.model.GetBucketWebsiteRequest;
import software.amazon.awssdk.services.s3.model.GetBucketWebsiteResponse;
import software.amazon.awssdk.services.s3.model.Grant;
import software.amazon.awssdk.services.s3.model.Grantee;
import software.amazon.awssdk.services.s3.model.LambdaFunctionConfiguration;
import software.amazon.awssdk.services.s3.model.LifecycleRule;
import software.amazon.awssdk.services.s3.model.NoncurrentVersionTransition;
import software.amazon.awssdk.services.s3.model.QueueConfiguration;
import software.amazon.awssdk.services.s3.model.Redirect;
import software.amazon.awssdk.services.s3.model.RedirectAllRequestsTo;
import software.amazon.awssdk.services.s3.model.ReplicationConfiguration;
import software.amazon.awssdk.services.s3.model.ReplicationRule;
import software.amazon.awssdk.services.s3.model.RoutingRule;
import software.amazon.awssdk.services.s3.model.S3Exception;
import software.amazon.awssdk.services.s3.model.ServerSideEncryptionByDefault;
import software.amazon.awssdk.services.s3.model.ServerSideEncryptionConfiguration;
import software.amazon.awssdk.services.s3.model.ServerSideEncryptionRule;
import software.amazon.awssdk.services.s3.model.SourceSelectionCriteria;
import software.amazon.awssdk.services.s3.model.SseKmsEncryptedObjects;
import software.amazon.awssdk.services.s3.model.TopicConfiguration;
import software.amazon.awssdk.services.s3.model.Transition;

public class AWSExporter
{
	private String executeTime;
	private String fileName;
	private XSSFWorkbook workbook;
	private XSSFHelper xssfHelper;
	private XSSFSheet vpcSheet;
	private XSSFSheet vpcPeeringSheet;
	private XSSFSheet subnetSheet;
	private XSSFSheet routeTableSheet;
	private XSSFSheet internetGatewaySheet;
	private XSSFSheet egressInternetGatewaySheet;
	private XSSFSheet natGatewaySheet;
	private XSSFSheet customerGatewaySheet;
	private XSSFSheet vpnGatewaySheet;
	private XSSFSheet vpnConnectionSheet;
	private XSSFSheet securityGroupSheet;
	private XSSFSheet ec2InstanceSheet;
	private XSSFSheet ebsSheet;
	private XSSFSheet autoScalingSheet;
	private XSSFSheet classicElbSheet;
	private XSSFSheet otherElbSheet;
	private XSSFSheet elastiCacheSheet;
	private XSSFSheet rdsClusterSheet;
	private XSSFSheet rdsInstanceSheet;
	private XSSFSheet kmsSheet;
	private XSSFSheet acmSheet;
	private XSSFSheet s3Sheet;
	private XSSFSheet directConnectSheet;
	private XSSFSheet directLocationSheet;
	private XSSFSheet virtualGatewaySheet;
	private XSSFSheet virtualInterfaceSheet;
	private XSSFSheet lagSheet;
	private XSSFSheet directConnectGatewaySheet;
	private XSSFSheet directorySheet;
	private AmazonAccess amazonAccess;
	private AmazonClients amazonClients;
	private Region region;
  
	private DateTimeFormatter formatDate = DateTimeFormatter
												.ofPattern("yyyy-MM-dd HH:mm:ss Z")
												.withZone(ZoneOffset.UTC);
	private DateFormat format = new SimpleDateFormat("yyyyMMddHHmmss");
  
	public AWSExporter(AmazonAccess amazonAccess, List<Region> regions, String profileName) {
		this.amazonAccess = amazonAccess;
		this.executeTime = format.format(new Date());
    
		System.out.println("AWS Exporter Start");
		for (Region region : regions) {
			make(region, profileName);
			write();
			try { this.workbook.close(); } catch (IOException e) { e.printStackTrace(); }
		}
	}
  
	private void write() {
		File dir = new File("exports");
		if (!dir.exists()) { dir.mkdir(); }
		File file = new File(dir, this.fileName);
		FileOutputStream os = null;
		try {
			System.out.println(this.fileName + " Write Start");
			os = new FileOutputStream(file);
			this.workbook.write(os);
			System.out.println(this.fileName + " Write End"); return;
		} catch (IOException e) {
			e.printStackTrace();
		} finally {
			if (os != null) { try { os.close(); } catch (Exception localException2) {} }
		}
	}
  
	private void initializeWorkbook() {
		this.workbook = new XSSFWorkbook();

		int sheetNumber = 1;

		this.vpcSheet = this.workbook.createSheet(sheetNumber++ + ".VPCs");
		this.vpcPeeringSheet = this.workbook.createSheet(sheetNumber++ + ".VPCPeerings");
		this.subnetSheet = this.workbook.createSheet(sheetNumber++ + ".Subnets");
		this.routeTableSheet = this.workbook.createSheet(sheetNumber++ + ".Route Tables");
		this.internetGatewaySheet = this.workbook.createSheet(sheetNumber++ + ".Internet Gateways");
		this.egressInternetGatewaySheet = this.workbook.createSheet(sheetNumber++ + ".Egress Only Internet Gateways");
		this.natGatewaySheet = this.workbook.createSheet(sheetNumber++ + ".NAT Gateways");
		this.customerGatewaySheet = this.workbook.createSheet(sheetNumber++ + ".Customer Gateways");
		this.vpnGatewaySheet = this.workbook.createSheet(sheetNumber++ + ".VPN Gateways");
		this.vpnConnectionSheet = this.workbook.createSheet(sheetNumber++ + ".VPN Connections");
		this.securityGroupSheet = this.workbook.createSheet(sheetNumber++ + ".Security Groups");
		this.ec2InstanceSheet = this.workbook.createSheet(sheetNumber++ + ".EC2 Instances");
		this.ebsSheet = this.workbook.createSheet(sheetNumber++ + ".EBS Volumes");
		this.autoScalingSheet = this.workbook.createSheet(sheetNumber++ + ".Auto Scaling Groups");
		this.classicElbSheet = this.workbook.createSheet(sheetNumber++ + ".Classic ELBs");
		this.otherElbSheet = this.workbook.createSheet(sheetNumber++ + ".Other ELBs");
		this.elastiCacheSheet = this.workbook.createSheet(sheetNumber++ + ".ElastiCaches");
		this.rdsClusterSheet = this.workbook.createSheet(sheetNumber++ + ".RDS-Clusters");
		this.rdsInstanceSheet = this.workbook.createSheet(sheetNumber++ + ".RDS-Instances");
		this.kmsSheet = this.workbook.createSheet(sheetNumber++ + ".KMS");
		this.acmSheet = this.workbook.createSheet(sheetNumber++ + ".ACM");
		this.s3Sheet = this.workbook.createSheet(sheetNumber++ + ".S3");
		this.directConnectSheet = this.workbook.createSheet(sheetNumber++ + ".Direct Connects");
		this.directLocationSheet = this.workbook.createSheet(sheetNumber++ + ".Direct Connect Locations");
		this.virtualGatewaySheet = this.workbook.createSheet(sheetNumber++ + ".Virtual Gateways");
		this.virtualInterfaceSheet = this.workbook.createSheet(sheetNumber++ + ".Virtual Interfaces");
		this.lagSheet = this.workbook.createSheet(sheetNumber++ + ".Lags");
		this.directConnectGatewaySheet = this.workbook.createSheet(sheetNumber++ + ".Direct Connect Gateways");
		this.directorySheet = this.workbook.createSheet(sheetNumber++ + ".Directory Services");
		
		this.xssfHelper = new XSSFHelper(this.workbook);
	}
  
	private String getNameTagValue(List<Tag> tags) {
		String name = "";
		for (Tag tag : tags) {
			if ("Name".equals(tag.key())) {
				name = tag.value();
				break;
			}
		}
		return name;
	}
  
	private String getAllTagDescriptionValue(List<software.amazon.awssdk.services.autoscaling.model.TagDescription> tagDescriptions) {
		StringBuffer tagValues = new StringBuffer();
		for(software.amazon.awssdk.services.autoscaling.model.TagDescription tagDescription : tagDescriptions) {
			if(tagValues.length() > 0) {
				tagValues.append("\n");
			}
			tagValues.append(tagDescription.key());
			tagValues.append("=");
			tagValues.append(tagDescription.value());
		}

		return tagValues.toString();
	}

	private String getAllTagValue(List<Tag> tags) {
		StringBuffer tagValues = new StringBuffer();
		for (Tag tag : tags) {
			if(tagValues.length() > 0) {
				tagValues.append("\n");
			}
			tagValues.append(tag.key());
			tagValues.append("=");
			tagValues.append(tag.value());
		}
		return tagValues.toString();
	}

	private static String rightPadding(String input, char ch, int L) {
		String result = String.format("%" + (-L) + "s", input).replace(' ', ch);
		return result;
	}

	private int progressPosition = 0;
	private void printProgress(String message) {
		this.progressPosition++;
		message = this.rightPadding(message, ' ', 50);
		this.printProgressBar(Long.parseLong(Long.toString(progressPosition*100/(this.workbook.getNumberOfSheets() + 1)*100)) / 100, message);
	}

	private void printProgressBar(long currentPosition, String message) {
		System.out.print(this.progressBar(100, currentPosition, 0, 100, message));
		System.out.print("\r");
		try {
			Thread.sleep(100);
		} catch (InterruptedException skip) {}
		if(currentPosition == 100) {
			System.out.println("\n");
		}
	}

	private String progressBar(int progressBarSize, long currentPosition, long startPositoin, long finishPosition, String message) {
		String bar = "";
		int nPositions = progressBarSize;
		char pb = '-'; //'?';
		char stat = '#'; //'?';
		for (int p = 0; p < nPositions; p++) {
			bar += pb;
		}
		int ststus = (int) (100 * (currentPosition - startPositoin) / (finishPosition - startPositoin));
		int move = (nPositions * ststus) / 100;
		return "|" + bar.substring(0, move).replace(pb, stat) + bar.substring(move, bar.length()) + "|" + ststus + "%| " + message;
	}

	private void make(Region region, String profileName) {
		this.region = region;
        
		this.fileName = ("[" + profileName + "][" + this.region.id() + "] AWSExport-" + this.executeTime + ".xlsx");
		System.out.println("Start " + this.region.id() + " Export. OutputFile [" + this.fileName + "]");
    
		initializeWorkbook();
		System.out.println("Excel Workbook Initialization Complete");
    
		this.amazonClients = new AmazonClients(this.amazonAccess, profileName, this.region);
		System.out.println("AWS Client Initialization Complete");
    
		makeSheetListHeader();
		System.out.println("Excel workbook sheet generation and subject line generation complete [" + this.workbook.getNumberOfSheets() + " ea Sheet]");

		this.progressPosition = 0;

		System.out.println("\n");

		try { makeSubnet(); printProgress("Subnets Export complete"); } catch (Exception e) { System.out.println("Subnets Export failure - [" + e.getMessage() + "]"); e.printStackTrace(); }
		
		HashMap<String, String> vpcMainRouteTables = null;
		try { vpcMainRouteTables = makeRouteTable(); printProgress("RouteTable Export complete"); } catch (Exception e) { System.out.println("RouteTable Export failure - [" + e.getMessage() + "]"); e.printStackTrace(); }
		try { makeVPC(vpcMainRouteTables); printProgress("VPC Export complete"); } catch (Exception e) { System.out.println("VPC Export failure - [" + e.getMessage() + "]"); e.printStackTrace(); }
		try { makeVPCPeering(); printProgress("VPC Export complete"); } catch (Exception e) { System.out.println("VPC Export failure - [" + e.getMessage() + "]"); e.printStackTrace(); }
		try { makeInternetGateway(); printProgress("InternetGateway Export complete"); } catch (Exception e) { System.out.println("InternetGateway Export failure - [" + e.getMessage() + "]"); e.printStackTrace(); }
		try { makeEgressInternetGateway(); printProgress("InternetGateway Export complete"); } catch (Exception e) { System.out.println("InternetGateway Export failure - [" + e.getMessage() + "]"); e.printStackTrace(); }
		try { makeNATGateway(); printProgress("EgressOnlyInternetGateway Export complete"); } catch (Exception e) { System.out.println("EgressOnlyInternetGateway Export failure - [" + e.getMessage() + "]"); e.printStackTrace(); }
		try { makeCustomerGateway(); printProgress("CustomerGateway Export complete"); } catch (Exception e) { System.out.println("CustomerGateway Export failure - [" + e.getMessage() + "]"); e.printStackTrace(); }
		try { makeVPNGateway(); printProgress("VPNGateway Export complete"); } catch (Exception e) { System.out.println("VPNGateway Export failure - [" + e.getMessage() + "]"); e.printStackTrace(); }
		try { makeVPNConnection(); printProgress("VPNConnection Export complete"); } catch (Exception e) { System.out.println("VPNConnection Export failure - [" + e.getMessage() + "]"); e.printStackTrace(); }
		try { makeSecurityGroup(); printProgress("SecurityGroup Export complete"); } catch (Exception e) { System.out.println("SecurityGroup Export failure - [" + e.getMessage() + "]"); e.printStackTrace(); }
		try { makeEC2Instance(); printProgress("EC2Instances Export complete"); } catch (Exception e) { System.out.println("EC2Instances Export failure - [" + e.getMessage() + "]"); e.printStackTrace(); }
		try { makeEBS(); printProgress("EBS Export complete"); } catch (Exception e) { System.out.println("EBS Export failure - [" + e.getMessage() + "]"); e.printStackTrace(); }
		try { makeAutoScaling(); printProgress("Auto Scaling Groups Export complete"); } catch (Exception e) { System.out.println("Auto Scaling Groups Export failure - [" + e.getMessage() + "]"); e.printStackTrace(); }
		try { makeClassicELB(); printProgress("Classic ELB Export complete"); } catch (Exception e) { System.out.println("Classic ELB Export failure - [" + e.getMessage() + "]"); e.printStackTrace(); }
		try { makeOtherELB(); printProgress("Other ELB Export complete"); } catch (Exception e) { System.out.println("Other ELB Export failure - [" + e.getMessage() + "]"); e.printStackTrace(); }
		try { makeElasticCache(); printProgress("ElasticCache Export complete"); } catch (Exception e) { System.out.println("ElasticCache Export failure - [" + e.getMessage() + "]"); e.printStackTrace(); }
		try { makeRDSCluster(); printProgress("RDS Clusters Export complete"); } catch (Exception e) { System.out.println("RDS Clusters Export failure - [" + e.getMessage() + "]"); e.printStackTrace(); }
		try { makeRDSInstance(); printProgress("RDS Instances Export complete"); } catch (Exception e) { System.out.println("RDS Instances Export failure - [" + e.getMessage() + "]"); e.printStackTrace(); }
		try { makeKMS(); printProgress("KMS Export complete"); } catch (Exception e) { System.out.println("KMS Export failure - [" + e.getMessage() + "]"); e.printStackTrace(); }
		try { makeACM(); printProgress("ACM Export complete"); } catch (Exception e) { System.out.println("ACM Export failure - [" + e.getMessage() + "]"); e.printStackTrace(); }
		try { makeS3(); printProgress("S3 Export complete"); } catch (Exception e) { System.out.println("S3 Export failure - [" + e.getMessage() + "]"); e.printStackTrace(); }
		
		try { makeDirectConnection(); printProgress("DirectConnect Export complete"); } catch (Exception e) { System.out.println("DirectConnect Export failure - [" + e.getMessage() + "]"); e.printStackTrace(); }
		try { makeDirectLocation(); printProgress("DirectLocation Export complete"); } catch (Exception e) { System.out.println("DirectLocation Export failure - [" + e.getMessage() + "]"); e.printStackTrace(); }
		try { makeVirtualGateway(); printProgress("VirtualGateway Export complete"); } catch (Exception e) { System.out.println("VirtualGateway Export failure - [" + e.getMessage() + "]"); e.printStackTrace(); }
		try { makeVirtualInterface(); printProgress("VirtualInterface Export complete"); } catch (Exception e) { System.out.println("VirtualInterface Export failure - [" + e.getMessage() + "]"); e.printStackTrace(); }
		try { makeLag(); printProgress("Lag Export complete"); } catch (Exception e) { System.out.println("Lag Export failure - [" + e.getMessage() + "]"); e.printStackTrace(); }
		try { makeDirectConnectGateway(); printProgress("DirectConnect Gateway Export complete"); } catch (Exception e) { System.out.println("DirectConnect Gateway Export failure - [" + e.getMessage() + "]"); e.printStackTrace(); }
		
		try { makeDirectoryService(); printProgress("Directory Service Export complete"); } catch (Exception e) { System.out.println("Directory Service Export failure - [" + e.getMessage() + "]"); e.printStackTrace(); }

		printProgress("All resources are exported !");

		System.out.println("");

		this.xssfHelper.setAutoSizeColumn();
	}

	private void makeAutoScaling() {
		XSSFRow row = null;

		DescribeAutoScalingGroupsResponse describeAutoScalingGroupsResponse = this.amazonClients.autoScalingClient.describeAutoScalingGroups();
		List<AutoScalingGroup> autoScalingGroups = describeAutoScalingGroupsResponse.autoScalingGroups();
		for(AutoScalingGroup autoScalingGroup : autoScalingGroups) {

			row = this.xssfHelper.createRow(this.autoScalingSheet, 1);
			this.xssfHelper.setLeftThinCell(row, Integer.toString(row.getRowNum() - 1));
			this.xssfHelper.setCell(row, autoScalingGroup.autoScalingGroupName());
			this.xssfHelper.setCell(row, (autoScalingGroup.launchTemplate() == null ? "" : autoScalingGroup.launchTemplate().launchTemplateName() + "[" + autoScalingGroup.launchTemplate().launchTemplateId() + "]"));
			this.xssfHelper.setCell(row, autoScalingGroup.launchConfigurationName());
			this.xssfHelper.setCell(row, Integer.toString(autoScalingGroup.instances().size()));
			this.xssfHelper.setCell(row, autoScalingGroup.status());
			this.xssfHelper.setCell(row, autoScalingGroup.desiredCapacity().toString());
			this.xssfHelper.setCell(row, autoScalingGroup.minSize().toString());
			this.xssfHelper.setCell(row, autoScalingGroup.maxSize().toString());

			List<String> availabilityZones = autoScalingGroup.availabilityZones();
			StringBuffer availabilityZoneBuffer = new StringBuffer();
			for(String availabilityZone : availabilityZones) {
				availabilityZoneBuffer.append(availabilityZone);
				availabilityZoneBuffer.append(",");
			}
			this.xssfHelper.setCell(row, availabilityZoneBuffer.toString());
			this.xssfHelper.setCell(row, autoScalingGroup.defaultCooldown().toString());
			this.xssfHelper.setCell(row, autoScalingGroup.healthCheckGracePeriod().toString());
			this.xssfHelper.setRightThinCell(row, this.getAllTagDescriptionValue(autoScalingGroup.tags()));

			if(autoScalingGroup.instances().size() > 0) { 
				row = this.xssfHelper.createRow(this.autoScalingSheet, 1);
				this.xssfHelper.setSubHeadLeftThinCell(row, "Instances");
				this.xssfHelper.setSubHeadCell(row, "");
				this.xssfHelper.setSubHeadCell(row, "");
				this.xssfHelper.setSubHeadCell(row, "");
				this.xssfHelper.setSubHeadCell(row, "");
				this.xssfHelper.setSubHeadCell(row, "");
				this.xssfHelper.setSubHeadCell(row, "");
				this.xssfHelper.setSubHeadCell(row, "");
				this.xssfHelper.setSubHeadCell(row, "Instance Id");
				this.xssfHelper.setSubHeadCell(row, "Lifecycle State");
				this.xssfHelper.setSubHeadCell(row, "Launch Configuration Name");
				this.xssfHelper.setSubHeadCell(row, "Availability Zone");
				this.xssfHelper.setSubHeadRightThinCell(row, "Health Status");
			
				this.autoScalingSheet.addMergedRegion(new CellRangeAddress(row.getRowNum(), row.getRowNum(), 0, 7));

				int instanceIdx = 0;
				int instanceRow = row.getRowNum();
				for(software.amazon.awssdk.services.autoscaling.model.Instance instance : autoScalingGroup.instances()) {
					row = this.xssfHelper.createRow(this.autoScalingSheet, 1);
					if(instanceIdx == 0) {
						this.xssfHelper.setHeadLeftThinCell(row, "Instances");
					} else {
						this.xssfHelper.setCell(row, "");	
					}
					this.xssfHelper.setCell(row, "");
					this.xssfHelper.setCell(row, "");
					this.xssfHelper.setCell(row, "");
					this.xssfHelper.setCell(row, "");
					this.xssfHelper.setCell(row, "");
					this.xssfHelper.setCell(row, "");
					this.xssfHelper.setCell(row, "");
					this.xssfHelper.setCell(row, instance.instanceId());
					this.xssfHelper.setCell(row, instance.lifecycleState().toString());
					this.xssfHelper.setCell(row, instance.launchConfigurationName());
					this.xssfHelper.setCell(row, instance.availabilityZone());
					this.xssfHelper.setRightThinCell(row, instance.healthStatus());
					instanceIdx++;
				}
				
				this.autoScalingSheet.addMergedRegion(new CellRangeAddress(instanceRow + 1, row.getRowNum(), 0, 7));
			}

		}

	}

	private void makeDirectoryService() {
		XSSFRow row = null;
		DescribeDirectoriesResponse describeDirectoriesResponse = this.amazonClients.directoryClient.describeDirectories();
		List<DirectoryDescription> directoryDescriptions = describeDirectoriesResponse.directoryDescriptions();
		for (DirectoryDescription directoryDescription : directoryDescriptions) {
			row = this.xssfHelper.createRow(this.directorySheet, 1);
			this.xssfHelper.setLeftThinCell(row, Integer.toString(row.getRowNum() - 1));
			this.xssfHelper.setCell(row, directoryDescription.directoryId());
			this.xssfHelper.setCell(row, this.getEnumName(directoryDescription.type()));
			this.xssfHelper.setCell(row, directoryDescription.alias());
			this.xssfHelper.setCell(row, directoryDescription.name());
			this.xssfHelper.setCell(row, directoryDescription.shortName());
			this.xssfHelper.setCell(row, this.getEnumName(directoryDescription.size()));
			this.xssfHelper.setCell(row, directoryDescription.accessUrl());
			this.xssfHelper.setCell(row, directoryDescription.desiredNumberOfDomainControllers().toString());
			
			DirectoryVpcSettingsDescription directoryVpcSettingsDescription = directoryDescription.vpcSettings();
			this.xssfHelper.setCell(row, directoryVpcSettingsDescription == null ? "" : directoryVpcSettingsDescription.vpcId());
			this.xssfHelper.setCell(row, directoryVpcSettingsDescription == null ? "" : directoryVpcSettingsDescription.securityGroupId());
			
			StringBuffer vpcAvas = new StringBuffer();
			if(directoryVpcSettingsDescription != null) {
			List<String> availabilityZones = directoryVpcSettingsDescription.availabilityZones();
				for(String availabilityZone : availabilityZones) {
					if(vpcAvas.length() > 0) vpcAvas.append("\n");
					vpcAvas.append(availabilityZone);
				}
			}
			this.xssfHelper.setCell(row, vpcAvas.toString());
			
			StringBuffer vpcSubs = new StringBuffer();
			if(directoryVpcSettingsDescription != null) {
			List<String> subnetIds = directoryVpcSettingsDescription.subnetIds();
				for(String subnetId : subnetIds) {
					if(vpcSubs.length() > 0) vpcSubs.append("\n");
					vpcSubs.append(subnetId);
				}
			}
			this.xssfHelper.setCell(row, vpcSubs.toString());
			
			DirectoryConnectSettingsDescription directoryConnectSettingsDescription = directoryDescription.connectSettings();
			this.xssfHelper.setCell(row, directoryConnectSettingsDescription == null ? "" : directoryConnectSettingsDescription.customerUserName());
			this.xssfHelper.setCell(row, directoryConnectSettingsDescription == null ? "" : directoryConnectSettingsDescription.vpcId());
			this.xssfHelper.setCell(row, directoryConnectSettingsDescription == null ? "" : directoryConnectSettingsDescription.securityGroupId());
			
			StringBuffer csAvas = new StringBuffer();
			if(directoryConnectSettingsDescription != null) {
				List<String> csavailabilityZones = directoryConnectSettingsDescription.availabilityZones();
				for(String availabilityZone : csavailabilityZones) {
					if(csAvas.length() > 0) csAvas.append("\n");
					csAvas.append(availabilityZone);
				}
			}
			this.xssfHelper.setCell(row, csAvas.toString());
			
			StringBuffer cips = new StringBuffer();
			if(directoryConnectSettingsDescription != null) {
				List<String> connectIps = directoryConnectSettingsDescription.connectIps();
				for(String connectIp : connectIps) {
					if(cips.length() > 0) cips.append("\n");
					cips.append(connectIp);
				}
			}
			this.xssfHelper.setCell(row, cips.toString());
			
			StringBuffer csSubs = new StringBuffer();
			if(directoryConnectSettingsDescription != null) {
				List<String> csSubnetIds = directoryConnectSettingsDescription.subnetIds();
				for(String csSubnetId : csSubnetIds) {
					if(csSubs.length() > 0) csSubs.append("\n");
					csSubs.append(csSubnetId);
				}
			}			
			this.xssfHelper.setCell(row, csSubs.toString());
			
			this.xssfHelper.setCell(row, directoryConnectSettingsDescription == null ? "" : directoryDescription.description());
			
			StringBuffer dias = new StringBuffer();
			if(directoryConnectSettingsDescription != null) {
				List<String> dnsIpAddrs = directoryDescription.dnsIpAddrs();
				for(String dnsIpAddr : dnsIpAddrs) {
					if(dias.length() > 0) dias.append("\n");
					dias.append(dnsIpAddr);
				}
			}
			this.xssfHelper.setCell(row, dias.toString());
			this.xssfHelper.setCell(row, directoryConnectSettingsDescription == null ? "" : this.getEnumName(directoryDescription.edition()));
			this.xssfHelper.setCell(row, directoryConnectSettingsDescription == null ? "" : directoryDescription.ssoEnabled().toString());
			this.xssfHelper.setCell(row, directoryConnectSettingsDescription == null ? "" : this.getEnumName(directoryDescription.stage()));
			this.xssfHelper.setCell(row, directoryConnectSettingsDescription == null ? "" : directoryDescription.stageLastUpdatedDateTime().toString());
			this.xssfHelper.setCell(row, directoryConnectSettingsDescription == null ? "" : directoryDescription.stageReason());
			
			RadiusSettings radiusSettings = directoryConnectSettingsDescription == null ? null : directoryDescription.radiusSettings();
			this.xssfHelper.setCell(row, radiusSettings == null ? "" : this.getEnumName(radiusSettings.authenticationProtocol()));
			this.xssfHelper.setCell(row, radiusSettings == null ? "" : radiusSettings.displayLabel());
			this.xssfHelper.setCell(row, radiusSettings == null ? "" : radiusSettings.radiusPort().toString());
			this.xssfHelper.setCell(row, radiusSettings == null ? "" : radiusSettings.radiusRetries().toString());
			this.xssfHelper.setCell(row, radiusSettings == null ? "" : radiusSettings.radiusTimeout().toString());
			this.xssfHelper.setCell(row, radiusSettings == null ? "" : radiusSettings.sharedSecret());
			this.xssfHelper.setCell(row, radiusSettings == null ? "" : radiusSettings.useSameUsername().toString());
			
			StringBuffer rss = new StringBuffer();
			if(radiusSettings != null) {
				List<String> radiusServers = radiusSettings.radiusServers();
				for(String radiusServer : radiusServers) {
					if(rss.length() > 0) rss.append("\n");
					rss.append(radiusServer);
				}
			}
			this.xssfHelper.setCell(row, rss.toString());
			this.xssfHelper.setCell(row, this.getEnumName(directoryDescription.radiusStatus()));
			
			StringBuffer dcrs = new StringBuffer();
			DescribeDomainControllersResponse describeDomainControllersResponse = this.amazonClients.directoryClient.describeDomainControllers(DescribeDomainControllersRequest.builder().directoryId(directoryDescription.directoryId()).build());
			List<DomainController> domainControllers = describeDomainControllersResponse.domainControllers();
			for(DomainController domainController : domainControllers) {
				if(dcrs.length() > 0) dcrs.append("\n\n");
				dcrs.append("Domain Controller");
				dcrs.append("\nDomain Controller Id=");
				dcrs.append(domainController.domainControllerId());
				dcrs.append("\nDirectory Id=");
				dcrs.append(domainController.directoryId());
				dcrs.append("\nVpc Id=");
				dcrs.append(domainController.vpcId());
				dcrs.append("\nSubnet Id=");
				dcrs.append(domainController.subnetId());
				dcrs.append("\nAvailabilityZone=");
				dcrs.append(domainController.availabilityZone());
				dcrs.append("\nDNS Ip Address=");
				dcrs.append(domainController.dnsIpAddr());
				dcrs.append("\nLaunch Time=");
				dcrs.append(domainController.launchTime().toString());
				dcrs.append("\nStatus=");
				dcrs.append(this.getEnumName(domainController.status()));
				dcrs.append("\nStatus Last Updated Date Time=");
				dcrs.append(domainController.statusLastUpdatedDateTime().toString());
				dcrs.append("\nStatus Reason=");
				dcrs.append(domainController.statusReason());
			}
			this.xssfHelper.setCell(row, dcrs.toString());
			
			StringBuffer trs = new StringBuffer();
			DescribeTrustsResponse describeTrustsResponse = this.amazonClients.directoryClient.describeTrusts(DescribeTrustsRequest.builder().directoryId(directoryDescription.directoryId()).build());
			List<Trust> trusts = describeTrustsResponse.trusts();
			for(Trust trust : trusts) {
				if(trs.length() > 0) trs.append("\n\n");
				trs.append("Trusts");
				trs.append("\nDirectory Id=");
				trs.append(trust.directoryId());
				trs.append("\nRemote Domain Name=");
				trs.append(trust.remoteDomainName());
				trs.append("\nId=");
				trs.append(trust.trustId());
				trs.append("\nType=");
				trs.append(this.getEnumName(trust.trustType()));
				trs.append("\nDirection=");
				trs.append(this.getEnumName(trust.trustDirection()));
				trs.append("\nState=");
				trs.append(this.getEnumName(trust.trustState()));
				trs.append("\nState Reason=");
				trs.append(trust.trustStateReason());
				trs.append("\nCreated Date Time=");
				trs.append(trust.createdDateTime().toString());
				trs.append("\nLast Updated Date Time=");
				trs.append(trust.lastUpdatedDateTime().toString());
				trs.append("\nState Last Updated Date Time=");
				trs.append(trust.stateLastUpdatedDateTime().toString());
			}
			
			this.xssfHelper.setCell(row, trs.toString());
			
			StringBuffer cfrs = new StringBuffer();
			try {
				DescribeConditionalForwardersResponse describeConditionalForwardersResponse = this.amazonClients.directoryClient.describeConditionalForwarders(DescribeConditionalForwardersRequest.builder().directoryId(directoryDescription.directoryId()).build());
				List<ConditionalForwarder> conditionalForwarders = describeConditionalForwardersResponse.conditionalForwarders();
				for(ConditionalForwarder conditionalForwarder : conditionalForwarders) {
					if(cfrs.length() > 0) cfrs.append("\n\n");
					cfrs.append("Conditional Forwarder");
					cfrs.append("\nRemote Domain Name=");
					cfrs.append(conditionalForwarder.remoteDomainName());
					cfrs.append("\nReplication Scope=");
					cfrs.append(this.getEnumName(conditionalForwarder.replicationScope()));
					
					StringBuffer cdias = new StringBuffer();
					List<String> dnsIpAddrs = conditionalForwarder.dnsIpAddrs();
					for(String dnsIpAddr : dnsIpAddrs) {
						if(cdias.length() > 0) cdias.append(", ");
						cdias.append(dnsIpAddr);
					}
					cfrs.append("\nDNS Ip Addresses=");
					cfrs.append(cdias.toString());
				}
			} catch(InvalidParameterException skip) {}
			this.xssfHelper.setCell(row, cfrs.toString());
			this.xssfHelper.setRightThinCell(row, directoryDescription.launchTime().toString());
			
		}

	}
	
	private void makeDirectConnection() {
		XSSFRow row = null;
		DescribeConnectionsResponse describeConnectionsResponse = this.amazonClients.directConnectClient.describeConnections();
		List<Connection> connections = describeConnectionsResponse.connections();
		for (Connection connection : connections) {
			row = this.xssfHelper.createRow(this.directConnectSheet, 1);
			this.xssfHelper.setLeftThinCell(row, Integer.toString(row.getRowNum() - 1));
			this.xssfHelper.setCell(row, connection.region());
			this.xssfHelper.setCell(row, connection.ownerAccount());
			this.xssfHelper.setCell(row, connection.connectionId());
			this.xssfHelper.setCell(row, connection.connectionName());
			this.xssfHelper.setCell(row, this.getEnumName(connection.connectionState()));
			this.xssfHelper.setCell(row, connection.location());
			this.xssfHelper.setCell(row, connection.awsDevice());
			this.xssfHelper.setCell(row, connection.vlan() == null ? "" : connection.vlan().toString());
			this.xssfHelper.setCell(row, connection.bandwidth());
			this.xssfHelper.setCell(row, connection.partnerName());
			this.xssfHelper.setCell(row, connection.lagId());
			this.xssfHelper.setRightThinCell(row, connection.loaIssueTime().toString());
		}
		
	}
	
	private void makeDirectLocation() {
		XSSFRow row = null;
		DescribeLocationsResponse describeLocationsResponse = this.amazonClients.directConnectClient.describeLocations();
		List<Location> locations = describeLocationsResponse.locations();
		for (Location location : locations) {
			row = this.xssfHelper.createRow(this.directLocationSheet, 1);
			this.xssfHelper.setLeftThinCell(row, Integer.toString(row.getRowNum() - 1));
			this.xssfHelper.setCell(row, this.region.id());
			this.xssfHelper.setCell(row, location.locationCode());
			this.xssfHelper.setRightThinCell(row, location.locationName());
		}
		
	}
	
	private void makeVirtualGateway() {
		XSSFRow row = null;
		DescribeVirtualGatewaysResponse describeVirtualGatewaysResponse = this.amazonClients.directConnectClient.describeVirtualGateways();
		List<VirtualGateway> virtualGateways = describeVirtualGatewaysResponse.virtualGateways();
		for (VirtualGateway virtualGateway : virtualGateways) {
			row = this.xssfHelper.createRow(this.virtualGatewaySheet, 1);
			this.xssfHelper.setLeftThinCell(row, Integer.toString(row.getRowNum() - 1));
			this.xssfHelper.setCell(row, this.region.id());
			this.xssfHelper.setCell(row, virtualGateway.virtualGatewayId());
			this.xssfHelper.setRightThinCell(row, virtualGateway.virtualGatewayState());
		}
		
	}
	
	private void makeVirtualInterface() {
		XSSFRow row = null;
		DescribeVirtualInterfacesResponse describeVirtualInterfacesResponse = this.amazonClients.directConnectClient.describeVirtualInterfaces();
		List<VirtualInterface> virtualInterfaces = describeVirtualInterfacesResponse.virtualInterfaces();
		for (VirtualInterface virtualInterface : virtualInterfaces) {
			row = this.xssfHelper.createRow(this.virtualInterfaceSheet, 1);
			this.xssfHelper.setLeftThinCell(row, Integer.toString(row.getRowNum() - 1));
			this.xssfHelper.setCell(row, this.region.id());
			this.xssfHelper.setCell(row, virtualInterface.virtualGatewayId());
			this.xssfHelper.setCell(row, virtualInterface.virtualInterfaceId());
			this.xssfHelper.setCell(row, virtualInterface.virtualInterfaceName());
			this.xssfHelper.setCell(row, this.getEnumName(virtualInterface.virtualInterfaceState()));
			this.xssfHelper.setCell(row, virtualInterface.virtualInterfaceType());
			this.xssfHelper.setCell(row, virtualInterface.asn().toString());
			this.xssfHelper.setCell(row, this.getEnumName(virtualInterface.addressFamily()));
			this.xssfHelper.setCell(row, virtualInterface.amazonSideAsn().toString());
			this.xssfHelper.setCell(row, virtualInterface.amazonAddress());
			this.xssfHelper.setCell(row, virtualInterface.authKey());
			
			StringBuffer bgps = new StringBuffer();
			List<BGPPeer> bgpPeers = virtualInterface.bgpPeers();
			for(BGPPeer bgpPeer : bgpPeers) {
				if(bgps.length() > 0) bgps.append("\n\n");
				bgps.append("ASN=");
				bgps.append(bgpPeer.asn());
				bgps.append("\nAddress Family=");
				bgps.append(this.getEnumName(bgpPeer.addressFamily()));
				bgps.append("\nAmazon Address=");
				bgps.append(bgpPeer.amazonAddress());
				bgps.append("\nCustomer Address=");
				bgps.append(bgpPeer.customerAddress());
				bgps.append("\nAuth Key=");
				bgps.append(bgpPeer.authKey());
				bgps.append("\nBGP Peer State=");
				bgps.append(this.getEnumName(bgpPeer.bgpPeerState()));
				bgps.append("\nBGP Status=");
				bgps.append(this.getEnumName(bgpPeer.bgpStatus()));
			}
			this.xssfHelper.setCell(row, bgps.toString());
			this.xssfHelper.setCell(row, virtualInterface.connectionId());
			this.xssfHelper.setCell(row, virtualInterface.customerAddress());
			this.xssfHelper.setCell(row, virtualInterface.customerRouterConfig());
			this.xssfHelper.setCell(row, virtualInterface.directConnectGatewayId());
			this.xssfHelper.setCell(row, virtualInterface.location());
			this.xssfHelper.setCell(row, virtualInterface.ownerAccount());
			
			StringBuffer rfps = new StringBuffer();
			List<RouteFilterPrefix> routeFilterPrefixs = virtualInterface.routeFilterPrefixes();
			for(RouteFilterPrefix routeFilterPrefix : routeFilterPrefixs) {
				if(rfps.length() > 0) rfps.append("\n");
				rfps.append(routeFilterPrefix.cidr());
			}
			this.xssfHelper.setCell(row, rfps.toString());
			this.xssfHelper.setRightThinCell(row, virtualInterface.vlan().toString());
		}
		
	}
	
	private void makeLag() {
		XSSFRow row = null;
		DescribeLagsResponse describeLagsResponse = this.amazonClients.directConnectClient.describeLags(DescribeLagsRequest.builder().build());
		List<Lag> lags = describeLagsResponse.lags();
		for (Lag lag : lags) {
			row = this.xssfHelper.createRow(this.lagSheet, 1);
			this.xssfHelper.setLeftThinCell(row, Integer.toString(row.getRowNum() - 1));
			this.xssfHelper.setCell(row, lag.region());
			this.xssfHelper.setCell(row, lag.allowsHostedConnections().toString());
			this.xssfHelper.setCell(row, lag.awsDevice());
			
			StringBuffer cons = new StringBuffer();
			List<Connection> connections = lag.connections();
			for(Connection connection : connections) {
				if(cons.length() > 0) cons.append("\n");
				cons.append(connection.connectionId());
				cons.append("|");
				cons.append(connection.connectionName());
				cons.append(" (");
				cons.append(connection.connectionState());
				cons.append(")");
			}
			this.xssfHelper.setCell(row, cons.toString());
			this.xssfHelper.setCell(row, lag.connectionsBandwidth());
			this.xssfHelper.setCell(row, lag.lagId());
			this.xssfHelper.setCell(row, lag.lagName());
			this.xssfHelper.setCell(row, this.getEnumName(lag.lagState()));
			this.xssfHelper.setCell(row, lag.minimumLinks().toString());
			this.xssfHelper.setCell(row, lag.numberOfConnections().toString());
			this.xssfHelper.setRightThinCell(row, lag.ownerAccount());			
		}
		
	}
	
	private void makeDirectConnectGateway() {
		XSSFRow row = null;
		DescribeDirectConnectGatewaysResponse describeDirectConnectGatewaysResponse = this.amazonClients.directConnectClient.describeDirectConnectGateways(DescribeDirectConnectGatewaysRequest.builder().build());
		List<DirectConnectGateway> directConnectGateways = describeDirectConnectGatewaysResponse.directConnectGateways();
		for (DirectConnectGateway directConnectGateway : directConnectGateways) {
			row = this.xssfHelper.createRow(this.directConnectGatewaySheet, 1);
			this.xssfHelper.setLeftThinCell(row, Integer.toString(row.getRowNum() - 1));
			this.xssfHelper.setCell(row, this.region.id());
			this.xssfHelper.setCell(row, directConnectGateway.directConnectGatewayId());
			this.xssfHelper.setCell(row, directConnectGateway.directConnectGatewayName());
			this.xssfHelper.setCell(row, directConnectGateway.amazonSideAsn().toString());
			this.xssfHelper.setCell(row, this.getEnumName(directConnectGateway.directConnectGatewayState()));
			this.xssfHelper.setCell(row, directConnectGateway.ownerAccount());
			this.xssfHelper.setRightThinCell(row, directConnectGateway.stateChangeError());
		}
	}
	
	private void makeEgressInternetGateway() {
		XSSFRow row = null;
		
		DescribeEgressOnlyInternetGatewaysResponse describeEgressOnlyInternetGatewaysResponse = this.amazonClients.ec2Client.describeEgressOnlyInternetGateways(DescribeEgressOnlyInternetGatewaysRequest.builder().build());
		List<EgressOnlyInternetGateway> egressOnlyInternetGateways = describeEgressOnlyInternetGatewaysResponse.egressOnlyInternetGateways();
		for (EgressOnlyInternetGateway egressOnlyInternetGateway : egressOnlyInternetGateways) {

			row = this.xssfHelper.createRow(this.egressInternetGatewaySheet, 1);
			this.xssfHelper.setLeftThinCell(row, Integer.toString(row.getRowNum() - 1));
			this.xssfHelper.setCell(row, this.region.id());
			this.xssfHelper.setCell(row, egressOnlyInternetGateway.egressOnlyInternetGatewayId());
			
			StringBuffer iga = new StringBuffer();
			List<InternetGatewayAttachment> internetGatewayAttachments = egressOnlyInternetGateway.attachments();
			for(InternetGatewayAttachment internetGatewayAttachment : internetGatewayAttachments) {
				if(iga.length() > 0) iga.append("\n");
				iga.append(internetGatewayAttachment.vpcId());
				iga.append("(");
				iga.append(internetGatewayAttachment.state());
				iga.append(")");
			}
			this.xssfHelper.setRightThinCell(row, iga.toString());
			
		}
		
	}
	
	private void makeNATGateway() {
		XSSFRow row = null;
		
		DescribeNatGatewaysResponse describeNatGatewaysResponse = this.amazonClients.ec2Client.describeNatGateways(DescribeNatGatewaysRequest.builder().build());
		List<NatGateway> natGateways = describeNatGatewaysResponse.natGateways();
		for(NatGateway natGateway : natGateways) {
			row = this.xssfHelper.createRow(this.natGatewaySheet, 1);
			this.xssfHelper.setLeftThinCell(row, Integer.toString(row.getRowNum() - 1));
			this.xssfHelper.setCell(row, this.region.id());
			this.xssfHelper.setCell(row, natGateway.natGatewayId());
			
			StringBuffer ngas = new StringBuffer();
			List<NatGatewayAddress> natGatewayAddresses = natGateway.natGatewayAddresses();
			for(NatGatewayAddress natGatewayAddress : natGatewayAddresses) {
				if(ngas.length() > 0) ngas.append("\n\n");
				ngas.append("\nNetworkInterfaceId=");
				ngas.append(natGatewayAddress.networkInterfaceId());
				ngas.append("\nPrivate Ip=");
				ngas.append(natGatewayAddress.privateIp());
				ngas.append("\nPublic Ip=");
				ngas.append(natGatewayAddress.publicIp());
				ngas.append("\nAllocation Id=");
				ngas.append(natGatewayAddress.allocationId());
			}
			this.xssfHelper.setCell(row, ngas.toString());
			this.xssfHelper.setCell(row, natGateway.vpcId());
			this.xssfHelper.setCell(row, natGateway.subnetId());
			this.xssfHelper.setCell(row, natGateway.failureCode());
			this.xssfHelper.setCell(row, natGateway.failureMessage());
			ProvisionedBandwidth provisionedBandwidth = natGateway.provisionedBandwidth();
			this.xssfHelper.setCell(row, provisionedBandwidth == null ? "" : provisionedBandwidth.provisioned());
			this.xssfHelper.setCell(row, provisionedBandwidth == null ? "" : provisionedBandwidth.provisionTime().toString());
			this.xssfHelper.setCell(row, provisionedBandwidth == null ? "" : provisionedBandwidth.requested());
			this.xssfHelper.setCell(row, provisionedBandwidth == null ? "" : provisionedBandwidth.requestTime().toString());
			this.xssfHelper.setCell(row, provisionedBandwidth == null ? "" : provisionedBandwidth.status());
			this.xssfHelper.setCell(row, this.getEnumName(natGateway.state()));
			this.xssfHelper.setCell(row, natGateway.createTime().toString());
			this.xssfHelper.setCell(row, natGateway.deleteTime() == null ? "" : natGateway.deleteTime().toString());
			this.xssfHelper.setRightThinCell(row, this.getAllTagValue(natGateway.tags()));
		}
		
	}
	
	private void makeCustomerGateway() {
		XSSFRow row = null;
		
		DescribeCustomerGatewaysResponse describeCustomerGatewaysResponse = this.amazonClients.ec2Client.describeCustomerGateways();
		List<CustomerGateway> customerGateways = describeCustomerGatewaysResponse.customerGateways();
		for(CustomerGateway customerGateway : customerGateways) {
			row = this.xssfHelper.createRow(this.customerGatewaySheet, 1);
			this.xssfHelper.setLeftThinCell(row, Integer.toString(row.getRowNum() - 1));
			this.xssfHelper.setCell(row, this.region.id());
			this.xssfHelper.setCell(row, customerGateway.customerGatewayId());
			this.xssfHelper.setCell(row, customerGateway.type());
			this.xssfHelper.setCell(row, customerGateway.ipAddress());
			this.xssfHelper.setCell(row, customerGateway.bgpAsn());
			this.xssfHelper.setCell(row, customerGateway.state());
			this.xssfHelper.setRightThinCell(row, this.getAllTagValue(customerGateway.tags()));
		}
		
	}
	
	private void makeVPNGateway() {
		XSSFRow row = null;
		
		DescribeVpnGatewaysResponse describeVpnGatewaysResponse = this.amazonClients.ec2Client.describeVpnGateways();
		List<VpnGateway> vpnGateways = describeVpnGatewaysResponse.vpnGateways();
		for(VpnGateway vpnGateway : vpnGateways) {
			row = this.xssfHelper.createRow(this.vpnGatewaySheet, 1);
			this.xssfHelper.setLeftThinCell(row, Integer.toString(row.getRowNum() - 1));
			this.xssfHelper.setCell(row, this.region.id());
			this.xssfHelper.setCell(row, vpnGateway.vpnGatewayId());
			this.xssfHelper.setCell(row, vpnGateway.availabilityZone());
			this.xssfHelper.setCell(row, this.getEnumName(vpnGateway.type()));
			this.xssfHelper.setCell(row, vpnGateway.amazonSideAsn().toString());
			this.xssfHelper.setCell(row, this.getEnumName(vpnGateway.state()));
			
			StringBuffer vas = new StringBuffer();
			List<VpcAttachment> vpcAttachments = vpnGateway.vpcAttachments();
			for(VpcAttachment vpcAttachment : vpcAttachments) {
				if(vas.length() > 0) vas.append("\n");
				vas.append(vpcAttachment.vpcId());
				vas.append(" (");
				vas.append(vpcAttachment.state());
				vas.append(")");
			}
			this.xssfHelper.setCell(row, vas.toString());
			this.xssfHelper.setRightThinCell(row, this.getAllTagValue(vpnGateway.tags()));
		}
		
	}
	
	private void makeVPNConnection() {
		XSSFRow row = null;
		
		DescribeVpnConnectionsResponse describeVpnConnectionsResponse = this.amazonClients.ec2Client.describeVpnConnections();
		List<VpnConnection> vpnConnections = describeVpnConnectionsResponse.vpnConnections();
		for(VpnConnection vpnConnection : vpnConnections) {
			row = this.xssfHelper.createRow(this.vpnConnectionSheet, 1);
			this.xssfHelper.setLeftThinCell(row, Integer.toString(row.getRowNum() - 1));
			this.xssfHelper.setCell(row, this.region.id());
			this.xssfHelper.setCell(row, vpnConnection.vpnConnectionId());
			this.xssfHelper.setCell(row, vpnConnection.vpnGatewayId());
			this.xssfHelper.setCell(row, vpnConnection.customerGatewayId());
			this.xssfHelper.setCell(row, vpnConnection.customerGatewayConfiguration());
			
			StringBuffer vgts = new StringBuffer();
			List<VgwTelemetry> vgwTelemetries = vpnConnection.vgwTelemetry();
			for(VgwTelemetry vgwTelemetry : vgwTelemetries) {
				if(vgts.length() > 0) vgts.append("\n\n");
				vgts.append("\nOutside IpAddress=");
				vgts.append(vgwTelemetry.outsideIpAddress());
				vgts.append("\nAccepted Route Count=");
				vgts.append(vgwTelemetry.acceptedRouteCount());
				vgts.append("\nLast Status Change=");
				vgts.append(vgwTelemetry.lastStatusChange().toString());
				vgts.append("\nStatus=");
				vgts.append(this.getEnumName(vgwTelemetry.status()));
				vgts.append("\nStatus Message=");
				vgts.append(vgwTelemetry.statusMessage());
			}
			this.xssfHelper.setCell(row, vgts.toString());
			this.xssfHelper.setCell(row, this.getEnumName(vpnConnection.type()));
			this.xssfHelper.setCell(row, vpnConnection.category());
			this.xssfHelper.setCell(row, vpnConnection.options().staticRoutesOnly().toString());
			
			StringBuffer vsrs = new StringBuffer();
			List<VpnStaticRoute> vpnStaticRoutes = vpnConnection.routes();
			for(VpnStaticRoute vpnStaticRoute : vpnStaticRoutes) {
				if(vsrs.length() > 0) vsrs.append("\n\n");
				vsrs.append("\nSource=");
				vsrs.append(vpnStaticRoute.source());
				vsrs.append("\nDestination=");
				vsrs.append(vpnStaticRoute.destinationCidrBlock());
				vsrs.append("\nState=");
				vsrs.append(this.getEnumName(vpnStaticRoute.state()));
			}
			this.xssfHelper.setCell(row, vsrs.toString());
			this.xssfHelper.setCell(row, this.getEnumName(vpnConnection.state()));
			this.xssfHelper.setRightThinCell(row, this.getAllTagValue(vpnConnection.tags()));
		}
		
	}
	
	
	private void makeVPCPeering() {
	
		XSSFRow row = null;
		
		DescribeVpcPeeringConnectionsResponse describeVpcPeeringConnectionsResponse = this.amazonClients.ec2Client.describeVpcPeeringConnections();
		List<VpcPeeringConnection> vpcPeeringConnections = describeVpcPeeringConnectionsResponse.vpcPeeringConnections();
		for (VpcPeeringConnection vpcPeeringConnection : vpcPeeringConnections) {
			
			row = this.xssfHelper.createRow(this.vpcPeeringSheet, 1);
			this.xssfHelper.setLeftThinCell(row, Integer.toString(row.getRowNum() - 1));
			this.xssfHelper.setCell(row, this.region.id());
			
			this.xssfHelper.setCell(row, vpcPeeringConnection.vpcPeeringConnectionId());
			this.getVpcPeeringVPCInfo(row, vpcPeeringConnection.requesterVpcInfo());						
			this.getVpcPeeringVPCInfo(row, vpcPeeringConnection.accepterVpcInfo());
			this.xssfHelper.setCell(row, vpcPeeringConnection.expirationTime() == null ? "" : vpcPeeringConnection.expirationTime().toString());
			VpcPeeringConnectionStateReason vpcPeeringConnectionStateReason = vpcPeeringConnection.status();
			this.xssfHelper.setCell(row, vpcPeeringConnectionStateReason == null ? "" : this.getEnumName(vpcPeeringConnectionStateReason.code()));
			this.xssfHelper.setCell(row, vpcPeeringConnectionStateReason == null ? "" : vpcPeeringConnectionStateReason.message());
			this.xssfHelper.setCell(row, this.getAllTagValue(vpcPeeringConnection.tags()));
		}
	}
	
	private void getVpcPeeringVPCInfo(XSSFRow row, VpcPeeringConnectionVpcInfo vpcPeeringConnectionVpcInfo) {
		this.xssfHelper.setCell(row, vpcPeeringConnectionVpcInfo.vpcId());
		this.xssfHelper.setCell(row, vpcPeeringConnectionVpcInfo.region());
		this.xssfHelper.setCell(row, vpcPeeringConnectionVpcInfo.cidrBlock());
		
		StringBuffer cidrs = new StringBuffer();
		List<CidrBlock> cidrBlocks = vpcPeeringConnectionVpcInfo.cidrBlockSet();
		for(CidrBlock cidrBlock : cidrBlocks) {
			if(cidrs.length() > 0) cidrs.append("\n");
			cidrs.append(cidrBlock.cidrBlock());
		}
		this.xssfHelper.setCell(row, cidrs.toString());
		
		StringBuffer ipv6s = new StringBuffer();
		List<Ipv6CidrBlock> ipv6CidrBlocks = vpcPeeringConnectionVpcInfo.ipv6CidrBlockSet();
		for(Ipv6CidrBlock ipv6CidrBlock : ipv6CidrBlocks) {
			if(ipv6s.length() > 0) ipv6s.append("\n");
			ipv6s.append(ipv6CidrBlock.ipv6CidrBlock());
		}
		this.xssfHelper.setCell(row, ipv6s.toString());
		this.xssfHelper.setCell(row, vpcPeeringConnectionVpcInfo.ownerId());
		VpcPeeringConnectionOptionsDescription vpcPeeringConnectionOptionsDescription = vpcPeeringConnectionVpcInfo.peeringOptions();
		this.xssfHelper.setCell(row, vpcPeeringConnectionOptionsDescription == null ? "" : vpcPeeringConnectionOptionsDescription.allowDnsResolutionFromRemoteVpc() == null ? "" : vpcPeeringConnectionOptionsDescription.allowDnsResolutionFromRemoteVpc().toString());
		this.xssfHelper.setCell(row, vpcPeeringConnectionOptionsDescription == null ? "" : vpcPeeringConnectionOptionsDescription.allowEgressFromLocalClassicLinkToRemoteVpc() == null ? "" : vpcPeeringConnectionOptionsDescription.allowEgressFromLocalClassicLinkToRemoteVpc().toString());
		this.xssfHelper.setRightThinCell(row, vpcPeeringConnectionOptionsDescription == null ? "" : vpcPeeringConnectionOptionsDescription.allowEgressFromLocalVpcToRemoteClassicLink() == null ? "" : vpcPeeringConnectionOptionsDescription.allowEgressFromLocalVpcToRemoteClassicLink().toString());
	}
	
	private void makeVPC(HashMap<String, String> vpcMainRouteTables) {
	    XSSFRow row = null;
	    List<Vpc> vpcs = this.amazonClients.ec2Client.describeVpcs().vpcs();
	    for (int i = 0; i < vpcs.size(); i++)
	    {
			Vpc vpc = vpcs.get(i);
			  
			row = this.xssfHelper.createRow(this.vpcSheet, 1);
			this.xssfHelper.setLeftThinCell(row, Integer.toString(i + 1));
			this.xssfHelper.setCell(row, this.region.id());
			this.xssfHelper.setCell(row, getNameTagValue(vpc.tags()));
			this.xssfHelper.setCell(row, vpc.vpcId());

			StringBuffer cidrs = new StringBuffer();
			for(VpcCidrBlockAssociation vpcCidrBlockAssociation : vpc.cidrBlockAssociationSet()) {
				if(cidrs.length() > 0) cidrs.append("\n");
				cidrs.append(vpcCidrBlockAssociation.cidrBlock());
			}
			this.xssfHelper.setCell(row, cidrs.toString());
			this.xssfHelper.setCell(row, vpc.dhcpOptionsId());
			this.xssfHelper.setCell(row, this.getEnumName(vpc.instanceTenancy()));
			
			if (vpcMainRouteTables != null) {
			    this.xssfHelper.setCell(row, vpcMainRouteTables.get(vpc.vpcId()));
			} else {
			    this.xssfHelper.setCell(row, "");
			}
			
			DescribeNetworkAclsResponse describeNetworkAclsResponse = this.amazonClients.ec2Client.describeNetworkAcls(
																				DescribeNetworkAclsRequest.builder().filters(
																							Filter.builder().name("vpc-id").values(vpc.vpcId()).build()
																							, Filter.builder().name("default").values("true").build()
																				).build()
																	  );
			NetworkAcl networkAcl = describeNetworkAclsResponse.networkAcls().get(0);
			this.xssfHelper.setCell(row, networkAcl.networkAclId() + " | " + getNameTagValue(networkAcl.tags()));
			
			this.xssfHelper.setCell(row, Boolean.toString(vpc.isDefault()));
			this.xssfHelper.setCell(row, this.getEnumName(vpc.state()));
			 
			StringBuffer vpcCidr = new StringBuffer();
			List<VpcCidrBlockAssociation> vpcCidrBlockAssociations = vpc.cidrBlockAssociationSet();
			for(VpcCidrBlockAssociation vpcCidrBlockAssociation : vpcCidrBlockAssociations) {
				  vpcCidr.append("AssociationId : " + vpcCidrBlockAssociation.associationId());
				  vpcCidr.append("\n");
				  vpcCidr.append("CIDR Block : " + vpcCidrBlockAssociation.cidrBlock());
				  VpcCidrBlockState vpcCidrBlockState = vpcCidrBlockAssociation.cidrBlockState();
				  if(vpcCidrBlockState != null) {
					  vpcCidr.append("\n");
					  vpcCidr.append("State : " + vpcCidrBlockState.state());
					  vpcCidr.append("\n");
					  vpcCidr.append("Status Message : " + vpcCidrBlockState.statusMessage());
				  }
				  vpcCidr.append("\n");
			}
			this.xssfHelper.setCell(row, vpcCidr.toString());
			  
			StringBuffer vpcIpv6Cidr = new StringBuffer();
			List<VpcIpv6CidrBlockAssociation> vpcIpv6CidrBlockAssociations = vpc.ipv6CidrBlockAssociationSet();
			for(VpcIpv6CidrBlockAssociation vpcIpv6CidrBlockAssociation : vpcIpv6CidrBlockAssociations) {
				vpcIpv6Cidr.append("AssociationId : " + vpcIpv6CidrBlockAssociation.associationId());
				vpcIpv6Cidr.append("\n");
				vpcIpv6Cidr.append("CIDR Block : " + vpcIpv6CidrBlockAssociation.ipv6CidrBlock());
				VpcCidrBlockState vpcCidrBlockState = vpcIpv6CidrBlockAssociation.ipv6CidrBlockState();
				if(vpcCidrBlockState != null) {
					vpcIpv6Cidr.append("\n");
					vpcIpv6Cidr.append("State : " + this.getEnumName(vpcCidrBlockState.state()));
					vpcIpv6Cidr.append("\n");
					vpcIpv6Cidr.append("Status Message : " + vpcCidrBlockState.statusMessage());
				}
			}
			this.xssfHelper.setCell(row, vpcIpv6Cidr.toString());
			  
			this.xssfHelper.setRightThinCell(row, this.getAllTagValue(vpc.tags()));
	      
	    }
	}
  
	private void makeSheetListHeader() {
		XSSFRow row = this.xssfHelper.createRow(this.vpcSheet, 1);
		this.xssfHelper.setHeadLeftThinCell(row, "No.");
		this.xssfHelper.setHeadCell(row, "Region");
		this.xssfHelper.setHeadCell(row, "Name");
		this.xssfHelper.setHeadCell(row, "VPC ID");
		this.xssfHelper.setHeadCell(row, "CIDR");
		this.xssfHelper.setHeadCell(row, "DHCP Options ID");
		this.xssfHelper.setHeadCell(row, "Instance Tenancy");
		this.xssfHelper.setHeadCell(row, "Main RouteTable");
		this.xssfHelper.setHeadCell(row, "Default ACL");
		this.xssfHelper.setHeadCell(row, "is Default");
		this.xssfHelper.setHeadCell(row, "State");
		this.xssfHelper.setHeadCell(row, "Cidr Block Associations");
		this.xssfHelper.setHeadCell(row, "Ipv6Cidr Block Associations");
		this.xssfHelper.setHeadRightThinCell(row, "Tags");
    
		row = this.xssfHelper.createRow(this.vpcPeeringSheet, 1);
		this.xssfHelper.setHeadLeftThinCell(row, "No.");
		this.xssfHelper.setHeadCell(row, "Region");
		this.xssfHelper.setHeadCell(row, "Vpc Peering Connection Id");
		this.xssfHelper.setHeadCell(row, "Requester Vpc Id");
		this.xssfHelper.setHeadCell(row, "Requester Region");
		this.xssfHelper.setHeadCell(row, "Requester CidrBlock");
		this.xssfHelper.setHeadCell(row, "Requester CidrBlock Set");
		this.xssfHelper.setHeadCell(row, "Requester Ipv6 CidrBlock Set");
		this.xssfHelper.setHeadCell(row, "Requester Owner Id");
		this.xssfHelper.setHeadCell(row, "Requester Allow Dns Resolution From Remote VPC");
		this.xssfHelper.setHeadCell(row, "Requester Allow Egress From Local ClassicLink To Remote Vpc");
		this.xssfHelper.setHeadCell(row, "Requester Allow Egress From Local Vpc To Remote ClassicLink");
		this.xssfHelper.setHeadCell(row, "Accepter Vpc Id");
		this.xssfHelper.setHeadCell(row, "Accepter Region");
		this.xssfHelper.setHeadCell(row, "Accepter CidrBlock");
		this.xssfHelper.setHeadCell(row, "Accepter CidrBlock Set");
		this.xssfHelper.setHeadCell(row, "Accepter Ipv6 CidrBlock Set");
		this.xssfHelper.setHeadCell(row, "Accepter Owner Id");
		this.xssfHelper.setHeadCell(row, "Accepter Allow Dns Resolution From Remote VPC");
		this.xssfHelper.setHeadCell(row, "Accepter Allow Egress From Local ClassicLink To Remote Vpc");
		this.xssfHelper.setHeadCell(row, "Accepter Allow Egress From Local Vpc To Remote ClassicLink");
		this.xssfHelper.setHeadCell(row, "Expiration Time");
		this.xssfHelper.setHeadCell(row, "State Reason Code");
		this.xssfHelper.setHeadCell(row, "State Reason Message");
		this.xssfHelper.setHeadRightThinCell(row, "Tags");
		
		row = this.xssfHelper.createRow(this.subnetSheet, 1);
		this.xssfHelper.setHeadLeftThinCell(row, "No.");
		this.xssfHelper.setHeadCell(row, "Region");
		this.xssfHelper.setHeadCell(row, "VPC ID");
		this.xssfHelper.setHeadCell(row, "Name");
		this.xssfHelper.setHeadCell(row, "SUBNET ID");
		this.xssfHelper.setHeadCell(row, "CIDR");
		this.xssfHelper.setHeadCell(row, "Availability Zone");
		this.xssfHelper.setHeadCell(row, "Route Table");
		this.xssfHelper.setHeadCell(row, "NetworkACL");
		this.xssfHelper.setHeadCell(row, "Assign Ipv6 Address On Creation");
		this.xssfHelper.setHeadCell(row, "Available IpAddress Count");
		this.xssfHelper.setHeadCell(row, "Default For AZ");
		this.xssfHelper.setHeadCell(row, "Ip V6 Cidr Block Association");
		this.xssfHelper.setHeadCell(row, "Map Public Ip On Launch");
		this.xssfHelper.setHeadCell(row, "State");
		this.xssfHelper.setHeadCell(row, "Tags");
		this.xssfHelper.setHeadRightThinCell(row, "Remark");
    
		row = this.xssfHelper.createRow(this.routeTableSheet, 1);
		this.xssfHelper.setHeadLeftThinCell(row, "No.");
		this.xssfHelper.setHeadCell(row, "Region");
		this.xssfHelper.setHeadCell(row, "VPC ID");
		this.xssfHelper.setHeadCell(row, "Name");
		this.xssfHelper.setHeadCell(row, "Route Table ID");
		this.xssfHelper.setHeadCell(row, "Main");
		this.xssfHelper.setHeadCell(row, "");
		this.xssfHelper.setHeadCell(row, "Tags");
		this.xssfHelper.setHeadRightThinCell(row, "");
		
		this.routeTableSheet.addMergedRegion(new CellRangeAddress(row.getRowNum(), row.getRowNum(), 5, 6));
		this.routeTableSheet.addMergedRegion(new CellRangeAddress(row.getRowNum(), row.getRowNum(), 7, 8));
    
		row = this.xssfHelper.createRow(this.internetGatewaySheet, 1);
		this.xssfHelper.setHeadLeftThinCell(row, "No.");
		this.xssfHelper.setHeadCell(row, "Region");
		this.xssfHelper.setHeadCell(row, "VPC ID");
		this.xssfHelper.setHeadCell(row, "Name");
		this.xssfHelper.setHeadCell(row, "InternetGateway ID");
		this.xssfHelper.setHeadRightThinCell(row, "Tags");
		
		row = this.xssfHelper.createRow(this.egressInternetGatewaySheet, 1);
		this.xssfHelper.setHeadLeftThinCell(row, "No.");
		this.xssfHelper.setHeadCell(row, "Region");
		this.xssfHelper.setHeadCell(row, "Egress Only Internet Gateway Id");
		this.xssfHelper.setHeadRightThinCell(row, "Attachments");
		
		row = this.xssfHelper.createRow(this.natGatewaySheet, 1);
		this.xssfHelper.setHeadLeftThinCell(row, "No.");
		this.xssfHelper.setHeadCell(row, "Region");
		this.xssfHelper.setHeadCell(row, "NatGateway Id");
		this.xssfHelper.setHeadCell(row, "NatGateway Address");
		this.xssfHelper.setHeadCell(row, "VPC Id");
		this.xssfHelper.setHeadCell(row, "Subnet Id");
		this.xssfHelper.setHeadCell(row, "Failure Code");
		this.xssfHelper.setHeadCell(row, "Failure Message");
		this.xssfHelper.setHeadCell(row, "Provisioned");
		this.xssfHelper.setHeadCell(row, "Provision Time");
		this.xssfHelper.setHeadCell(row, "Provision Requested");
		this.xssfHelper.setHeadCell(row, "Provision Request Time");
		this.xssfHelper.setHeadCell(row, "Provision Status");
		this.xssfHelper.setHeadCell(row, "State");
		this.xssfHelper.setHeadCell(row, "Create Time");
		this.xssfHelper.setHeadCell(row, "Delete Time");
		this.xssfHelper.setHeadRightThinCell(row, "Tags");
				
		row = this.xssfHelper.createRow(this.customerGatewaySheet, 1);
		this.xssfHelper.setHeadLeftThinCell(row, "No.");
		this.xssfHelper.setHeadCell(row, "Region");
		this.xssfHelper.setHeadCell(row, "Customer Gateway Id");
		this.xssfHelper.setHeadCell(row, "Type");
		this.xssfHelper.setHeadCell(row, "IpAddress");
		this.xssfHelper.setHeadCell(row, "BGP ASN");
		this.xssfHelper.setHeadCell(row, "State");
		this.xssfHelper.setHeadRightThinCell(row, "Tags");
				
		row = this.xssfHelper.createRow(this.vpnGatewaySheet, 1);
		this.xssfHelper.setHeadLeftThinCell(row, "No.");
		this.xssfHelper.setHeadCell(row, "Region");
		this.xssfHelper.setHeadCell(row, "VPN Gateway Id");
		this.xssfHelper.setHeadCell(row, "Availability Zone");
		this.xssfHelper.setHeadCell(row, "Type");
		this.xssfHelper.setHeadCell(row, "Amazon Side ASN");
		this.xssfHelper.setHeadCell(row, "State");
		this.xssfHelper.setHeadCell(row, "VPC Attachments");
		this.xssfHelper.setHeadRightThinCell(row, "Tags");

		row = this.xssfHelper.createRow(this.vpnConnectionSheet, 1);
		this.xssfHelper.setHeadLeftThinCell(row, "No.");
		this.xssfHelper.setHeadCell(row, "Region");
		this.xssfHelper.setHeadCell(row, "VPN Connection Id");
		this.xssfHelper.setHeadCell(row, "VPN Gateway Id");
		this.xssfHelper.setHeadCell(row, "Customer Gateway Id");
		this.xssfHelper.setHeadCell(row, "Customer Gateway Configuration");
		this.xssfHelper.setHeadCell(row, "VGW Telemetry");
		this.xssfHelper.setHeadCell(row, "Type");
		this.xssfHelper.setHeadCell(row, "Category");
		this.xssfHelper.setHeadCell(row, "Static Routes Only");
		this.xssfHelper.setHeadCell(row, "VPN Static Routes");
		this.xssfHelper.setHeadCell(row, "State");
		this.xssfHelper.setHeadRightThinCell(row, "Tags");
		
		row = this.xssfHelper.createRow(this.securityGroupSheet, 1);
		this.xssfHelper.setHeadLeftThinCell(row, "No.");
		this.xssfHelper.setHeadCell(row, "Region");
		this.xssfHelper.setHeadCell(row, "VPC ID");
		this.xssfHelper.setHeadCell(row, "Group ID [Group Name] - Description");
		this.xssfHelper.setHeadCell(row, "");
		this.xssfHelper.setHeadCell(row, "");
		this.xssfHelper.setHeadCell(row, "");
		this.xssfHelper.setHeadCell(row, "");
		this.xssfHelper.setHeadCell(row, "Tags");
		this.xssfHelper.setHeadCell(row, "");
		this.xssfHelper.setHeadCell(row, "");
		this.xssfHelper.setHeadCell(row, "");
		this.xssfHelper.setHeadRightThinCell(row, "");
    
		this.securityGroupSheet.addMergedRegion(new CellRangeAddress(row.getRowNum(), row.getRowNum(), 3, 7));
		this.securityGroupSheet.addMergedRegion(new CellRangeAddress(row.getRowNum(), row.getRowNum(), 8, 12));
    
		row = this.xssfHelper.createRow(this.ec2InstanceSheet, 1);
		this.xssfHelper.setHeadLeftThinCell(row, "No.");
		this.xssfHelper.setHeadCell(row, "Instance Name");
		this.xssfHelper.setHeadCell(row, "Instance ID");
		this.xssfHelper.setHeadCell(row, "Type");
		this.xssfHelper.setHeadCell(row, "Zone");
		this.xssfHelper.setHeadCell(row, "Private Ip");
		this.xssfHelper.setHeadCell(row, "Private DNS Name");
		this.xssfHelper.setHeadCell(row, "EIP - Public Ip");
		this.xssfHelper.setHeadCell(row, "Public DNS Name");
		this.xssfHelper.setHeadCell(row, "Security Groups");
		this.xssfHelper.setHeadCell(row, "Launch Time");
		this.xssfHelper.setHeadCell(row, "InstanceState");
		this.xssfHelper.setHeadCell(row, "State Reason");
		this.xssfHelper.setHeadCell(row, "State Transition Reason");
		this.xssfHelper.setHeadCell(row, "Monitoring State");
		this.xssfHelper.setHeadCell(row, "Source Destination Check");
		this.xssfHelper.setHeadCell(row, "Spot Instance Request ID");
		this.xssfHelper.setHeadCell(row, "Sriov Net Support");
		this.xssfHelper.setHeadCell(row, "Virtualization Type");
		this.xssfHelper.setHeadCell(row, "Platform");
		this.xssfHelper.setHeadCell(row, "Architecture");
		this.xssfHelper.setHeadCell(row, "KernelId");
		this.xssfHelper.setHeadCell(row, "EnaSupport");
		this.xssfHelper.setHeadCell(row, "HyperVisor");
		this.xssfHelper.setHeadCell(row, "ClientToken");
		this.xssfHelper.setHeadCell(row, "Ami Launch Index");
		this.xssfHelper.setHeadCell(row, "ImageId");
		this.xssfHelper.setHeadCell(row, "Instance Lifecycle");
		this.xssfHelper.setHeadCell(row, "Key Name");
		this.xssfHelper.setHeadCell(row, "CPU");
		this.xssfHelper.setHeadCell(row, "Elastic Gpu Associations");
		this.xssfHelper.setHeadCell(row, "EBS Optimized");
		this.xssfHelper.setHeadCell(row, "RAM Disk ID");
		this.xssfHelper.setHeadCell(row, "Root Device Name");
		this.xssfHelper.setHeadCell(row, "Root Device Type");
		this.xssfHelper.setHeadCell(row, "Block Device Mappings");
		this.xssfHelper.setHeadCell(row, "Network Interfaces");
		this.xssfHelper.setHeadCell(row, "Placement");
		this.xssfHelper.setHeadCell(row, "Product Codes");
		this.xssfHelper.setHeadCell(row, "IAM Instance Profile");
		this.xssfHelper.setHeadRightThinCell(row, "Tags");
    
		row = this.xssfHelper.createRow(this.ebsSheet, 1);
		this.xssfHelper.setHeadLeftThinCell(row, "No.");
		this.xssfHelper.setHeadCell(row, "Region");
		this.xssfHelper.setHeadCell(row, "Name");
		this.xssfHelper.setHeadCell(row, "Volume ID");
		this.xssfHelper.setHeadCell(row, "Volume Type");
		this.xssfHelper.setHeadCell(row, "Size");
		this.xssfHelper.setHeadCell(row, "IOPS");
		this.xssfHelper.setHeadCell(row, "Availability Zone");
		this.xssfHelper.setHeadCell(row, "Snapshot ID");
		this.xssfHelper.setHeadCell(row, "Attachment Information");
		this.xssfHelper.setHeadCell(row, "Encrypted");
		this.xssfHelper.setHeadCell(row, "KMS ID");
		this.xssfHelper.setHeadRightThinCell(row, "Tags");

		row = this.xssfHelper.createRow(this.autoScalingSheet, 1);
		this.xssfHelper.setHeadLeftThinCell(row, "No.");
		this.xssfHelper.setHeadCell(row, "Name");
		this.xssfHelper.setHeadCell(row, "Launch template");
		this.xssfHelper.setHeadCell(row, "Launch Configuration");
		this.xssfHelper.setHeadCell(row, "Instances");
		this.xssfHelper.setHeadCell(row, "Status");
		this.xssfHelper.setHeadCell(row, "Desired");
		this.xssfHelper.setHeadCell(row, "Min");
		this.xssfHelper.setHeadCell(row, "Max");
		this.xssfHelper.setHeadCell(row, "Availability Zones");
		this.xssfHelper.setHeadCell(row, "Default Cooldown");
		this.xssfHelper.setHeadCell(row, "Health Check Grace period");
		this.xssfHelper.setHeadRightThinCell(row, "Tags");
    
		row = this.xssfHelper.createRow(this.classicElbSheet, 1);
		this.xssfHelper.setHeadLeftThinCell(row, "No.");
		this.xssfHelper.setHeadCell(row, "Region");
		this.xssfHelper.setHeadCell(row, "VPC ID");
		this.xssfHelper.setHeadCell(row, "ELB Type");
		this.xssfHelper.setHeadCell(row, "Availability Zone");
		this.xssfHelper.setHeadCell(row, "ELB Name");
		this.xssfHelper.setHeadCell(row, "DNS Name");
		this.xssfHelper.setHeadCell(row, "Instances");
		this.xssfHelper.setHeadCell(row, "Listeners");
		this.xssfHelper.setHeadCell(row, "");
		this.xssfHelper.setHeadCell(row, "");
		this.xssfHelper.setHeadCell(row, "");
		this.xssfHelper.setHeadRightThinCell(row, "Health Check");
    
		this.classicElbSheet.addMergedRegion(new CellRangeAddress(row.getRowNum(), row.getRowNum(), 8, 11));
        
		row = this.xssfHelper.createRow(this.otherElbSheet, 1);
		this.xssfHelper.setHeadLeftThinCell(row, "No.");
		this.xssfHelper.setHeadCell(row, "Region");
		this.xssfHelper.setHeadCell(row, "VPC ID");
		this.xssfHelper.setHeadCell(row, "Availability Zone");
		this.xssfHelper.setHeadCell(row, "ELB Type");
		this.xssfHelper.setHeadCell(row, "Scheme");
		this.xssfHelper.setHeadCell(row, "IpAddress Type");
		this.xssfHelper.setHeadCell(row, "ELB Name");
		this.xssfHelper.setHeadCell(row, "Canonical Hosted Zone Id");
		this.xssfHelper.setHeadCell(row, "Status");
		this.xssfHelper.setHeadCell(row, "SecurityGroups");
		this.xssfHelper.setHeadCell(row, "DNS Name");
		this.xssfHelper.setHeadCell(row, "");
		this.xssfHelper.setHeadCell(row, "");
		this.xssfHelper.setHeadCell(row, "ELB ARN");
		this.xssfHelper.setHeadRightThinCell(row, "Tags");
    
		this.otherElbSheet.addMergedRegion(new CellRangeAddress(row.getRowNum(), row.getRowNum(), 11, 13));
    
		row = this.xssfHelper.createRow(this.elastiCacheSheet, 1);
		this.xssfHelper.setHeadLeftThinCell(row, "No.");
		this.xssfHelper.setHeadCell(row, "Region");
		this.xssfHelper.setHeadCell(row, "Cluster Name");
		this.xssfHelper.setHeadCell(row, "Engine");
		this.xssfHelper.setHeadCell(row, "Engine Version Compatibility");
		this.xssfHelper.setHeadCell(row, "Primary Endpoint");
		this.xssfHelper.setHeadCell(row, "Parameter Group");
		this.xssfHelper.setHeadCell(row, "Subnet Group");
		this.xssfHelper.setHeadCell(row, "Security Groups");
		this.xssfHelper.setHeadCell(row, "Maintenance Window");
		this.xssfHelper.setHeadCell(row, "Nodes");
		this.xssfHelper.setHeadCell(row, "");
		this.xssfHelper.setHeadCell(row, "");
		this.xssfHelper.setHeadCell(row, "");
		this.xssfHelper.setHeadRightThinCell(row, "");
    
		row = this.xssfHelper.createRow(this.elastiCacheSheet, 1);
		this.xssfHelper.setHeadLeftThinCell(row, "");
		this.xssfHelper.setHeadCell(row, "");
		this.xssfHelper.setHeadCell(row, "");
		this.xssfHelper.setHeadCell(row, "");
		this.xssfHelper.setHeadCell(row, "");
		this.xssfHelper.setHeadCell(row, "");
		this.xssfHelper.setHeadCell(row, "");
		this.xssfHelper.setHeadCell(row, "");
		this.xssfHelper.setHeadCell(row, "");
		this.xssfHelper.setHeadCell(row, "");
		this.xssfHelper.setHeadCell(row, "Node Name");
		this.xssfHelper.setHeadCell(row, "Node Availability Zone");
		this.xssfHelper.setHeadCell(row, "Current Role");
		this.xssfHelper.setHeadCell(row, "Node Type");
		this.xssfHelper.setHeadRightThinCell(row, "Endpoint");
    
		this.elastiCacheSheet.addMergedRegion(new CellRangeAddress(row.getRowNum() - 1, row.getRowNum(), 0, 0));
		this.elastiCacheSheet.addMergedRegion(new CellRangeAddress(row.getRowNum() - 1, row.getRowNum(), 1, 1));
		this.elastiCacheSheet.addMergedRegion(new CellRangeAddress(row.getRowNum() - 1, row.getRowNum(), 2, 2));
		this.elastiCacheSheet.addMergedRegion(new CellRangeAddress(row.getRowNum() - 1, row.getRowNum(), 3, 3));
		this.elastiCacheSheet.addMergedRegion(new CellRangeAddress(row.getRowNum() - 1, row.getRowNum(), 4, 4));
		this.elastiCacheSheet.addMergedRegion(new CellRangeAddress(row.getRowNum() - 1, row.getRowNum(), 5, 5));
		this.elastiCacheSheet.addMergedRegion(new CellRangeAddress(row.getRowNum() - 1, row.getRowNum(), 6, 6));
		this.elastiCacheSheet.addMergedRegion(new CellRangeAddress(row.getRowNum() - 1, row.getRowNum(), 7, 7));
		this.elastiCacheSheet.addMergedRegion(new CellRangeAddress(row.getRowNum() - 1, row.getRowNum(), 8, 8));
		this.elastiCacheSheet.addMergedRegion(new CellRangeAddress(row.getRowNum() - 1, row.getRowNum(), 9, 9));
		this.elastiCacheSheet.addMergedRegion(new CellRangeAddress(row.getRowNum() - 1, row.getRowNum() - 1, 10, 14));
		
		row = this.xssfHelper.createRow(this.rdsClusterSheet, 1);
		this.xssfHelper.setHeadLeftThinCell(row, "No.");
		this.xssfHelper.setHeadCell(row, "Kind");
		this.xssfHelper.setHeadCell(row, "DB Cluster Identifier");
		this.xssfHelper.setHeadCell(row, "Availability Zones");
		this.xssfHelper.setHeadCell(row, "DB Subnet Group");
		this.xssfHelper.setHeadCell(row, "DB Cluster Option Group");
		this.xssfHelper.setHeadCell(row, "DB Cluster Parameter Group");
		this.xssfHelper.setHeadCell(row, "DB Cluster Resource ID");
		this.xssfHelper.setHeadCell(row, "Database Name");
		this.xssfHelper.setHeadCell(row, "CharacterSet");
		this.xssfHelper.setHeadCell(row, "Allocated Storage");
		this.xssfHelper.setHeadCell(row, "Clone Group ID");
		this.xssfHelper.setHeadCell(row, "Backtrack Window");
		this.xssfHelper.setHeadCell(row, "Backup Retention Period");
		this.xssfHelper.setHeadCell(row, "Backtrack Consumed Change Records");
		this.xssfHelper.setHeadCell(row, "DB Cluster ARN");
		this.xssfHelper.setHeadCell(row, "DB Cluster Role");
		this.xssfHelper.setHeadCell(row, "Earliest Backtrack Time");
		this.xssfHelper.setHeadCell(row, "Earliest Restorable Time");
		this.xssfHelper.setHeadCell(row, "Enabled CloudwatchLogs Exports");
		this.xssfHelper.setHeadCell(row, "Endpoint");
		this.xssfHelper.setHeadCell(row, "Engine");
		this.xssfHelper.setHeadCell(row, "EngineVersion");
		this.xssfHelper.setHeadCell(row, "HostedZone Id");
		this.xssfHelper.setHeadCell(row, "IAM Database Authentication Enabled");
		this.xssfHelper.setHeadCell(row, "Kms Key Id");
		this.xssfHelper.setHeadCell(row, "Latest Restorable Time");
		this.xssfHelper.setHeadCell(row, "Master Username");
		this.xssfHelper.setHeadCell(row, "Multi AZ");
		this.xssfHelper.setHeadCell(row, "Percent Progress");
		this.xssfHelper.setHeadCell(row, "Port");
		this.xssfHelper.setHeadCell(row, "Preferred BackupWindow");
		this.xssfHelper.setHeadCell(row, "Preferred Maintenance Window");
		this.xssfHelper.setHeadCell(row, "Reader Endpoint");
		this.xssfHelper.setHeadCell(row, "Read Replica Identifiers");
		this.xssfHelper.setHeadCell(row, "Replication Source Identifier");
		this.xssfHelper.setHeadCell(row, "Status");
		this.xssfHelper.setHeadCell(row, "Storage Encrypted");
		this.xssfHelper.setHeadCell(row, "Vpc SecurityGroups");
		this.xssfHelper.setHeadRightThinCell(row, "DB Cluster Members");
		
		row = this.xssfHelper.createRow(this.rdsInstanceSheet, 1);
		this.xssfHelper.setHeadLeftThinCell(row, "No.");
		this.xssfHelper.setHeadCell(row, "Kind");
		this.xssfHelper.setHeadCell(row, "DB Cluster Identifier");
		this.xssfHelper.setHeadCell(row, "DB Instance Identifier");
		this.xssfHelper.setHeadCell(row, "DB Instance Class");
		this.xssfHelper.setHeadCell(row, "DB Instance Port");
		this.xssfHelper.setHeadCell(row, "DB Instance Status");
		this.xssfHelper.setHeadCell(row, "Availability Zones");
		this.xssfHelper.setHeadCell(row, "DB Subnet Group");
		this.xssfHelper.setHeadCell(row, "DB Security Groups");
		this.xssfHelper.setHeadCell(row, "DB Parameter Groups");
		this.xssfHelper.setHeadCell(row, "DB Name");
		this.xssfHelper.setHeadCell(row, "CharacterSet");
		this.xssfHelper.setHeadCell(row, "CA Certificateidentifier");
		this.xssfHelper.setHeadCell(row, "Auto Minor Version Upgrade");
		this.xssfHelper.setHeadCell(row, "Backup Retention Period");
		this.xssfHelper.setHeadCell(row, "Copy Tags To Snapshot");
		this.xssfHelper.setHeadCell(row, "DBI Resource ID");
		this.xssfHelper.setHeadCell(row, "Allocated Storage");
		this.xssfHelper.setHeadCell(row, "DB Instance ARN");
		this.xssfHelper.setHeadCell(row, "Domain Memberships");
		this.xssfHelper.setHeadCell(row, "Enabled Cloudwatch Logs Exports");
		this.xssfHelper.setHeadCell(row, "Endpoint Address");
		this.xssfHelper.setHeadCell(row, "Endpoint HostedZone Id");
		this.xssfHelper.setHeadCell(row, "Endpoint Port");
		this.xssfHelper.setHeadCell(row, "Engine");
		this.xssfHelper.setHeadCell(row, "Engine Version");
		this.xssfHelper.setHeadCell(row, "Enhanced Monitoring Resource ARN");
		this.xssfHelper.setHeadCell(row, "IAM Database Authentication Enabled");
		this.xssfHelper.setHeadCell(row, "Instance Create Time");
		this.xssfHelper.setHeadCell(row, "Iops");
		this.xssfHelper.setHeadCell(row, "Kms Key ID");
		this.xssfHelper.setHeadCell(row, "Lastest Restorable Time");
		this.xssfHelper.setHeadCell(row, "License Model");
		this.xssfHelper.setHeadCell(row, "Master Username");
		this.xssfHelper.setHeadCell(row, "Monitoring Interval");
		this.xssfHelper.setHeadCell(row, "Monitoring Role ARN");
		this.xssfHelper.setHeadCell(row, "Multi AZ");
		this.xssfHelper.setHeadCell(row, "Option Group Memberships");
		this.xssfHelper.setHeadCell(row, "Performance Insights Enabled");
		this.xssfHelper.setHeadCell(row, "Performance Insights KMS Key Id");
		this.xssfHelper.setHeadCell(row, "Performance Insights Retention Period");
		this.xssfHelper.setHeadCell(row, "Preferred Backup Window");
		this.xssfHelper.setHeadCell(row, "Preferred Maintenance Window");
		this.xssfHelper.setHeadCell(row, "Processor Features");
		this.xssfHelper.setHeadCell(row, "Promotion Tier");
		this.xssfHelper.setHeadCell(row, "Publicly Accessible");
		this.xssfHelper.setHeadCell(row, "Read Replica DB Cluster Identifiers");
		this.xssfHelper.setHeadCell(row, "Read Replica DB Instance Identifiers");
		this.xssfHelper.setHeadCell(row, "Read Replica Source DB Instance Identifier");
		this.xssfHelper.setHeadCell(row, "Secondary Availability Zone");
		this.xssfHelper.setHeadCell(row, "Status Infos");
		this.xssfHelper.setHeadCell(row, "Storage Encrypted");
		this.xssfHelper.setHeadCell(row, "Storage Type");
		this.xssfHelper.setHeadCell(row, "Tde Credential ARN");
		this.xssfHelper.setHeadCell(row, "Time Zone");
		this.xssfHelper.setHeadRightThinCell(row, "VPC Security Groups");
		
		row = this.xssfHelper.createRow(this.kmsSheet, 1);
		this.xssfHelper.setHeadLeftThinCell(row, "No.");
		this.xssfHelper.setHeadCell(row, "Key Manager");
		this.xssfHelper.setHeadCell(row, "Alias ARN");
		this.xssfHelper.setHeadCell(row, "Alias Name");
		this.xssfHelper.setHeadCell(row, "Key ID");
		this.xssfHelper.setHeadCell(row, "Key ARN");
		this.xssfHelper.setHeadCell(row, "Account ID");
		this.xssfHelper.setHeadCell(row, "Creation Date");
		this.xssfHelper.setHeadCell(row, "Deletion Date");
		this.xssfHelper.setHeadCell(row, "Description");
		this.xssfHelper.setHeadCell(row, "Enabled");
		this.xssfHelper.setHeadCell(row, "Expiration Model");
		this.xssfHelper.setHeadCell(row, "Key State");
		this.xssfHelper.setHeadCell(row, "Key Usage");
		this.xssfHelper.setHeadCell(row, "Origin");
		this.xssfHelper.setHeadRightThinCell(row, "Valid To");
		
		row = this.xssfHelper.createRow(this.acmSheet, 1);
		this.xssfHelper.setHeadLeftThinCell(row, "No.");
		this.xssfHelper.setHeadCell(row, "Certificate ARN");
		this.xssfHelper.setHeadCell(row, "Certificate Authority ARN");
		this.xssfHelper.setHeadCell(row, "Domain Name");
		this.xssfHelper.setHeadCell(row, "Status");
		this.xssfHelper.setHeadCell(row, "Subject");
		this.xssfHelper.setHeadCell(row, "Subject Alternative Names");
		this.xssfHelper.setHeadCell(row, "Type");
		this.xssfHelper.setHeadCell(row, "Key Algorithm");
		this.xssfHelper.setHeadCell(row, "Key Usages");
		this.xssfHelper.setHeadCell(row, "Serial");
		this.xssfHelper.setHeadCell(row, "Signature Algorithm");
		this.xssfHelper.setHeadCell(row, "Certificate Transparency Logging Preference");
		this.xssfHelper.setHeadCell(row, "Domain Validation Options");
		this.xssfHelper.setHeadCell(row, "Extended Key Usages");
		this.xssfHelper.setHeadCell(row, "Failure Reason");
		this.xssfHelper.setHeadCell(row, "Issuer");
		this.xssfHelper.setHeadCell(row, "Renewal Eligibility");
		this.xssfHelper.setHeadCell(row, "Renewal Status");
		this.xssfHelper.setHeadCell(row, "Renewal Domain Validation Options");
		this.xssfHelper.setHeadCell(row, "Created At");
		this.xssfHelper.setHeadCell(row, "Imported At");
		this.xssfHelper.setHeadCell(row, "InUse By");
		this.xssfHelper.setHeadCell(row, "Issued At");
		this.xssfHelper.setHeadCell(row, "Revoked At");
		this.xssfHelper.setHeadCell(row, "Revocation Reason");
		this.xssfHelper.setHeadCell(row, "Not After");
		this.xssfHelper.setHeadRightThinCell(row, "Not Before");
		
		row = this.xssfHelper.createRow(this.s3Sheet, 1);
		this.xssfHelper.setHeadLeftThinCell(row, "No.");
		this.xssfHelper.setHeadCell(row, "Owner");
		this.xssfHelper.setHeadCell(row, "Name");
		this.xssfHelper.setHeadCell(row, "Creation Date");
		this.xssfHelper.setHeadCell(row, "Bucket Location");
		this.xssfHelper.setHeadCell(row, "Server Side Encryption");
		this.xssfHelper.setHeadCell(row, "isAccelerateEnabled");
		this.xssfHelper.setHeadCell(row, "Accelerate Status");
		this.xssfHelper.setHeadCell(row, "Cross Origin Configuration");
		this.xssfHelper.setHeadCell(row, "Bucket Policy");
		this.xssfHelper.setHeadCell(row, "isLoggingEnabled");
		this.xssfHelper.setHeadCell(row, "Logging Destination BucketName");
		this.xssfHelper.setHeadCell(row, "LogFilePrefix");
		this.xssfHelper.setHeadCell(row, "isMfaDeleteEnabled");
		this.xssfHelper.setHeadCell(row, "Bucket Versioning Status");
		this.xssfHelper.setHeadCell(row, "AllTagSets");
		this.xssfHelper.setHeadCell(row, "TagSet");
		this.xssfHelper.setHeadCell(row, "Bucket Replication RoleARN");
		this.xssfHelper.setHeadCell(row, "Bucket Replication Rule");
		this.xssfHelper.setHeadCell(row, "Bucket Notification Configuration");
		this.xssfHelper.setHeadCell(row, "Bucket Website Configuration Error Document");
		this.xssfHelper.setHeadCell(row, "Bucket Website Configuration Index Document");
		this.xssfHelper.setHeadCell(row, "Bucket Website Redirect Rule Hostname");
		this.xssfHelper.setHeadCell(row, "Bucket Website Redirect Rule HttpRedirectCode");
		this.xssfHelper.setHeadCell(row, "Bucket Website Redirect Rule Protocol");
		this.xssfHelper.setHeadCell(row, "Bucket Website Redirect Rule ReplacekeyPrefixWidth");
		this.xssfHelper.setHeadCell(row, "Bucket Website Redirect Rule ReplacekeyWidth");
		this.xssfHelper.setHeadCell(row, "Bucket Website Routing Rule");
		this.xssfHelper.setHeadCell(row, "Access Control Owner");
		this.xssfHelper.setHeadCell(row, "Access Control Grants");
		this.xssfHelper.setHeadRightThinCell(row, "Bucket Lifecycle Configuration");
		
		row = this.xssfHelper.createRow(this.directConnectSheet, 1);
		this.xssfHelper.setHeadLeftThinCell(row, "No.");
		this.xssfHelper.setHeadCell(row, "Region");
		this.xssfHelper.setHeadCell(row, "Owner Account");
		this.xssfHelper.setHeadCell(row, "Connection Id");
		this.xssfHelper.setHeadCell(row, "Connection Name");
		this.xssfHelper.setHeadCell(row, "Connection State");
		this.xssfHelper.setHeadCell(row, "Location");
		this.xssfHelper.setHeadCell(row, "AWS Device");
		this.xssfHelper.setHeadCell(row, "VLAN");
		this.xssfHelper.setHeadCell(row, "Bandwidth");
		this.xssfHelper.setHeadCell(row, "Partner Name");
		this.xssfHelper.setHeadCell(row, "Lag ID");
		this.xssfHelper.setHeadRightThinCell(row, "Loa Issue Time");
		
		row = this.xssfHelper.createRow(this.directLocationSheet, 1);
		this.xssfHelper.setHeadLeftThinCell(row, "No.");
		this.xssfHelper.setHeadCell(row, "Region");
		this.xssfHelper.setHeadCell(row, "Location Code");
		this.xssfHelper.setHeadRightThinCell(row, "Location Name");
		
		row = this.xssfHelper.createRow(this.virtualGatewaySheet, 1);
		this.xssfHelper.setHeadLeftThinCell(row, "No.");
		this.xssfHelper.setHeadCell(row, "Region");
		this.xssfHelper.setHeadCell(row, "Virtual Gateway Id");
		this.xssfHelper.setHeadRightThinCell(row, "Virtual Gateway State");
		
		row = this.xssfHelper.createRow(this.virtualInterfaceSheet, 1);
		this.xssfHelper.setHeadLeftThinCell(row, "No.");
		this.xssfHelper.setHeadCell(row, "Region");
		this.xssfHelper.setHeadCell(row, "Virtual Gateway Id");
		this.xssfHelper.setHeadCell(row, "Virtual Interface Id");
		this.xssfHelper.setHeadCell(row, "Virtual Interface Name");
		this.xssfHelper.setHeadCell(row, "Virtual Interface State");
		this.xssfHelper.setHeadCell(row, "Virtual Interface Type");
		this.xssfHelper.setHeadCell(row, "ASN");
		this.xssfHelper.setHeadCell(row, "Address Family");
		this.xssfHelper.setHeadCell(row, "Amazon Side ASN");
		this.xssfHelper.setHeadCell(row, "Amazon Address");
		this.xssfHelper.setHeadCell(row, "Auth Key");
		this.xssfHelper.setHeadCell(row, "BGP Peers");
		this.xssfHelper.setHeadCell(row, "Connection ID");
		this.xssfHelper.setHeadCell(row, "Customer Address");
		this.xssfHelper.setHeadCell(row, "Customer Router Config");
		this.xssfHelper.setHeadCell(row, "Direct Connect Gateway ID");
		this.xssfHelper.setHeadCell(row, "Location");
		this.xssfHelper.setHeadCell(row, "Owner Account");
		this.xssfHelper.setHeadCell(row, "Router Filter Prefixes");
		this.xssfHelper.setHeadRightThinCell(row, "VLAN");
		
		row = this.xssfHelper.createRow(this.lagSheet, 1);
		this.xssfHelper.setHeadLeftThinCell(row, "No.");
		this.xssfHelper.setHeadCell(row, "Region");
		this.xssfHelper.setHeadCell(row, "Allows Hosted Connections");
		this.xssfHelper.setHeadCell(row, "AWS Device");
		this.xssfHelper.setHeadCell(row, "Connections");
		this.xssfHelper.setHeadCell(row, "Connections Bandwidth");
		this.xssfHelper.setHeadCell(row, "Lag Id");
		this.xssfHelper.setHeadCell(row, "Lag Name");
		this.xssfHelper.setHeadCell(row, "Lag State");
		this.xssfHelper.setHeadCell(row, "Minimum Links");
		this.xssfHelper.setHeadCell(row, "Number Of Connections");
		this.xssfHelper.setHeadRightThinCell(row, "Owner Account");
		
		row = this.xssfHelper.createRow(this.directConnectGatewaySheet, 1);
		this.xssfHelper.setHeadLeftThinCell(row, "No.");
		this.xssfHelper.setHeadCell(row, "Region");
		this.xssfHelper.setHeadCell(row, "Direct Connect Gateway Id");
		this.xssfHelper.setHeadCell(row, "Direct Connect Gateway Name");
		this.xssfHelper.setHeadCell(row, "Amazon Side ASN");
		this.xssfHelper.setHeadCell(row, "Direct Connect Gateway State");
		this.xssfHelper.setHeadCell(row, "Owner Account");
		this.xssfHelper.setHeadRightThinCell(row, "State Change Error");
		
		row = this.xssfHelper.createRow(this.directorySheet, 1);
		this.xssfHelper.setHeadLeftThinCell(row, "No.");
		this.xssfHelper.setHeadCell(row, "Directory Id");
		this.xssfHelper.setHeadCell(row, "Type");
		this.xssfHelper.setHeadCell(row, "Alias");
		this.xssfHelper.setHeadCell(row, "Name");
		this.xssfHelper.setHeadCell(row, "Short Name");
		this.xssfHelper.setHeadCell(row, "Size");
		this.xssfHelper.setHeadCell(row, "AccessUrl");
		this.xssfHelper.setHeadCell(row, "Desired Number Of Domain Controllers");
		this.xssfHelper.setHeadCell(row, "VpcSettings Vpc Id");
		this.xssfHelper.setHeadCell(row, "VpcSettings SecurityGroupId");
		this.xssfHelper.setHeadCell(row, "VpcSettings AvailabilityZones");
		this.xssfHelper.setHeadCell(row, "VpcSettings Subnet Ids");
		this.xssfHelper.setHeadCell(row, "ConnectSettings Customer UserName");
		this.xssfHelper.setHeadCell(row, "ConnectSettings Vpc Id");
		this.xssfHelper.setHeadCell(row, "ConnectSettings SecurityGroupId");
		this.xssfHelper.setHeadCell(row, "ConnectSettings AvalilabilityZones");
		this.xssfHelper.setHeadCell(row, "ConnectSettings Connect Ips");
		this.xssfHelper.setHeadCell(row, "ConnectSettings Subnet Ids");
		this.xssfHelper.setHeadCell(row, "Description");
		this.xssfHelper.setHeadCell(row, "DNS Ip Addresses");
		this.xssfHelper.setHeadCell(row, "Edition");
		this.xssfHelper.setHeadCell(row, "SSO Enabled");
		this.xssfHelper.setHeadCell(row, "Stage");
		this.xssfHelper.setHeadCell(row, "Stage Last updated Date Time");
		this.xssfHelper.setHeadCell(row, "Stage Reason");
		this.xssfHelper.setHeadCell(row, "Radius Authentication Protocol");
		this.xssfHelper.setHeadCell(row, "Radius Display Label");
		this.xssfHelper.setHeadCell(row, "Radius Port");
		this.xssfHelper.setHeadCell(row, "Radius Retries");
		this.xssfHelper.setHeadCell(row, "Radius Timeout");
		this.xssfHelper.setHeadCell(row, "Radius Shared Secret");
		this.xssfHelper.setHeadCell(row, "Radius Use Same Username");
		this.xssfHelper.setHeadCell(row, "Radius Servers");
		this.xssfHelper.setHeadCell(row, "Radius Status");
		this.xssfHelper.setHeadCell(row, "Domain Controllers");
		this.xssfHelper.setHeadCell(row, "Trusts");
		this.xssfHelper.setHeadCell(row, "Conditional Forwarders");
		this.xssfHelper.setHeadRightThinCell(row, "Launch Time");
	}
	
  
	private void makeS3() {
		
		XSSFRow row = null;
		
		List<Bucket> buckets = this.amazonClients.s3Client.listBuckets().buckets();
		for (Bucket bucket : buckets) {
			
			row = this.xssfHelper.createRow(this.s3Sheet, 1);			
			this.xssfHelper.setLeftThinCell(row, Integer.toString(row.getRowNum() - 1));
			this.xssfHelper.setCell(row, "");
			this.xssfHelper.setCell(row, bucket.name());
			this.xssfHelper.setCell(row, bucket.creationDate().toString());
			
			try {
				this.xssfHelper.setCell(row, this.amazonClients.s3Client.getBucketLocation(GetBucketLocationRequest.builder().bucket(bucket.name()).build()).locationConstraintAsString());

			} catch(S3Exception skip) {		
				this.xssfHelper.setCell(row, "");	
			} catch(Exception e) {
				this.xssfHelper.setCell(row, "");
				e.printStackTrace();
			}
			
			try {
				StringBuffer bets = new StringBuffer();
				GetBucketEncryptionResponse getBucketEncryptionResponse = this.amazonClients.s3Client.getBucketEncryption(GetBucketEncryptionRequest.builder().bucket(bucket.name()).build());
				ServerSideEncryptionConfiguration serverSideEncryptionConfiguration = getBucketEncryptionResponse.serverSideEncryptionConfiguration();
				List<ServerSideEncryptionRule> serverSideEncryptionRules = serverSideEncryptionConfiguration.rules();
				for(ServerSideEncryptionRule serverSideEncryptionRule : serverSideEncryptionRules) {
					ServerSideEncryptionByDefault serverSideEncryptionByDefault = serverSideEncryptionRule.applyServerSideEncryptionByDefault();
					
					if(bets.length() > 0) bets.append("\n");
					bets.append(serverSideEncryptionByDefault.kmsMasterKeyID());
					bets.append("|");
					bets.append(serverSideEncryptionByDefault.sseAlgorithm());
				}
				this.xssfHelper.setCell(row,  bets.toString());
			} catch(S3Exception skip) {		
				this.xssfHelper.setCell(row, "");		
			} catch(Exception e) {	
				this.xssfHelper.setCell(row, "");
				e.printStackTrace();
			}
			
			try {
				GetBucketAccelerateConfigurationResponse getBucketAccelerateConfigurationResponse = this.amazonClients.s3Client.getBucketAccelerateConfiguration(GetBucketAccelerateConfigurationRequest.builder().bucket(bucket.name()).build());
				this.xssfHelper.setCell(row,  "");
				this.xssfHelper.setCell(row,  getBucketAccelerateConfigurationResponse.statusAsString());
			} catch(S3Exception skip) {
				this.xssfHelper.setCell(row, "");
				this.xssfHelper.setCell(row, "");
			} catch(Exception e) {
				this.xssfHelper.setCell(row, "");
				this.xssfHelper.setCell(row, "");
				e.printStackTrace();
			}
			
			try {
				StringBuffer crrs = new StringBuffer();
				GetBucketCorsResponse getBucketCorsResponse = this.amazonClients.s3Client.getBucketCors(GetBucketCorsRequest.builder().bucket(bucket.name()).build());
				
				if(getBucketCorsResponse != null) {
					List<CORSRule> corsRules = getBucketCorsResponse.corsRules();
					for(CORSRule corsRule : corsRules) {
						
						if(crrs.length() > 0) crrs.append("\n");
						crrs.append("Id=");
						crrs.append("");
						crrs.append(", MaxAgeSeconds=");
						crrs.append(Integer.toString(corsRule.maxAgeSeconds()));
						
						crrs.append(", AllowedHeaders=");
						List<String> allowedHeaders = corsRule.allowedHeaders();
						for(String allowedHeader : allowedHeaders) {
							crrs.append(allowedHeader);
							crrs.append("|");
						}
						crrs.append(", AllowedMethods=");
						List<String> allowedMethods = corsRule.allowedMethods();
						for(String allowedMethod : allowedMethods) {
							crrs.append(allowedMethod);
							crrs.append("|");
						}
						crrs.append(", AllowedOrigins=");
						List<String> allowedOrigins = corsRule.allowedOrigins();
						for(String allowedOrigin : allowedOrigins) {
							crrs.append(allowedOrigin);
							crrs.append("|");
						}
						crrs.append(", ExposedHeaders=");
						List<String> exposedHeaders = corsRule.exposeHeaders();
						for(String exposedHeader : exposedHeaders) {
							crrs.append(exposedHeader);
							crrs.append("|");
						}
						
					}
				}
				
				this.xssfHelper.setCell(row, crrs.toString());

			} catch(S3Exception skip) {
				this.xssfHelper.setCell(row, "");
			} catch(Exception e) {
				this.xssfHelper.setCell(row, "");
				e.printStackTrace();
			}
			
			try {
				this.xssfHelper.setCell(row, this.amazonClients.s3Client.getBucketPolicy(GetBucketPolicyRequest.builder().bucket(bucket.name()).build()).policy());
			} catch(S3Exception skip) {
				this.xssfHelper.setCell(row, "");
			} catch(Exception e) {
				this.xssfHelper.setCell(row, "");
				e.printStackTrace();
			}
			
			try {
				GetBucketLoggingResponse getBucketLoggingResponse = this.amazonClients.s3Client.getBucketLogging(GetBucketLoggingRequest.builder().bucket(bucket.name()).build());
				this.xssfHelper.setCell(row, getBucketLoggingResponse.loggingEnabled() == null ? "false" : "true");
				this.xssfHelper.setCell(row, getBucketLoggingResponse.loggingEnabled() == null ? "" : getBucketLoggingResponse.loggingEnabled().targetBucket());
				this.xssfHelper.setCell(row, getBucketLoggingResponse.loggingEnabled() == null ? "" : getBucketLoggingResponse.loggingEnabled().targetPrefix());
			} catch(S3Exception skip) {
				this.xssfHelper.setCell(row, "");
				this.xssfHelper.setCell(row, "");
				this.xssfHelper.setCell(row, "");
			} catch(Exception e) {
				this.xssfHelper.setCell(row, "");
				this.xssfHelper.setCell(row, "");
				this.xssfHelper.setCell(row, "");
				e.printStackTrace();
			}
			
			try {
				GetBucketVersioningResponse getBucketVersioningResponse = this.amazonClients.s3Client.getBucketVersioning(GetBucketVersioningRequest.builder().bucket(bucket.name()).build());
				if(getBucketVersioningResponse != null) {
					this.xssfHelper.setCell(row, getBucketVersioningResponse.mfaDeleteAsString());
					this.xssfHelper.setCell(row, getBucketVersioningResponse.statusAsString());
				} else {
					this.xssfHelper.setCell(row, "");
					this.xssfHelper.setCell(row, "");
				}
			} catch(S3Exception skip) {
				this.xssfHelper.setCell(row, "");
				this.xssfHelper.setCell(row, "");
			} catch(Exception e) {
				this.xssfHelper.setCell(row, "");
				this.xssfHelper.setCell(row, "");
				e.printStackTrace();
			}
			
			try {
				StringBuffer btcs = new StringBuffer();
				GetBucketTaggingResponse getBucketTaggingResponse = this.amazonClients.s3Client.getBucketTagging(GetBucketTaggingRequest.builder().bucket(bucket.name()).build());
				
				if(getBucketTaggingResponse != null) {
					List<software.amazon.awssdk.services.s3.model.Tag> tagSets = getBucketTaggingResponse.tagSet();
					for(software.amazon.awssdk.services.s3.model.Tag tag : tagSets) {
						if(btcs.length() > 0) btcs.append("\n");
						btcs.append(tag.key());
						btcs.append("=");
						btcs.append(tag.value());
					}
				}
				this.xssfHelper.setCell(row, btcs.toString());
			} catch(S3Exception skip) {
				this.xssfHelper.setCell(row, "");
				this.xssfHelper.setCell(row, "");
			} catch(Exception e) {
				this.xssfHelper.setCell(row, "");
				this.xssfHelper.setCell(row, "");
				e.printStackTrace();
			}
			
			try {
				GetBucketReplicationResponse getBucketReplicationResponse = this.amazonClients.s3Client.getBucketReplication(GetBucketReplicationRequest.builder().bucket(bucket.name()).build());
				
				ReplicationConfiguration replicationConfiguration = getBucketReplicationResponse.replicationConfiguration();
				
				this.xssfHelper.setCell(row, replicationConfiguration.role());
				
				StringBuffer rrs = new StringBuffer();
				List<ReplicationRule> replicationRules = replicationConfiguration.rules();
				for(ReplicationRule replicationRule : replicationRules) {
					if(rrs.length() > 0) rrs.append("\n\n");
					
					rrs.append(replicationRule.id());

					Destination destination = replicationRule.destination();
					rrs.append("\nAccessControlTranslation=");
					rrs.append(destination.accessControlTranslation());
					rrs.append("\nAccount=");
					rrs.append(destination.account());
					rrs.append("\nBucketARN=");
					rrs.append(destination.bucket());
					
					rrs.append("\nReplicaKmsKeyID=");
					rrs.append(destination.encryptionConfiguration().replicaKmsKeyID());
					
					rrs.append("\nStorageClass=");
					rrs.append(destination.storageClass());
					rrs.append("\nPrefix=");
					rrs.append(replicationRule.prefix());
					SourceSelectionCriteria sourceSelectionCriteria = replicationRule.sourceSelectionCriteria();
					SseKmsEncryptedObjects sseKmsEncryptedObjects = sourceSelectionCriteria.sseKmsEncryptedObjects();
					rrs.append("\nSSE KMS Encrypted Objects Status=");
					rrs.append(this.getEnumName(sseKmsEncryptedObjects.status()));
					rrs.append("\nReplicationRule Status=");
					rrs.append(this.getEnumName(replicationRule.status()));
				}
				this.xssfHelper.setCell(row, rrs.toString());
			} catch(S3Exception skip) {
				this.xssfHelper.setCell(row, "");
				this.xssfHelper.setCell(row, "");
			} catch(Exception e) {
				this.xssfHelper.setCell(row, "");
				this.xssfHelper.setCell(row, "");
				e.printStackTrace();
			}
			
			try {			
				StringBuffer bncs = new StringBuffer();
				GetBucketNotificationConfigurationResponse getBucketNotificationConfigurationResponse = this.amazonClients.s3Client.getBucketNotificationConfiguration(GetBucketNotificationConfigurationRequest.builder().bucket(bucket.name()).build());
				
				bncs.append("Topic Configurations : ");
				bncs.append("\n");
				List<QueueConfiguration> queueConfigurations = getBucketNotificationConfigurationResponse.queueConfigurations();
				for (QueueConfiguration queueConfiguration : queueConfigurations) {
					bncs.append("Queue ARN = ");
					bncs.append(queueConfiguration.queueArn());
					bncs.append("\n");
					bncs.append("Event = ");
					bncs.append(queueConfiguration.eventsAsStrings());
				}
				bncs.append("\n\n");
				bncs.append("Queue Configurations : ");;
				bncs.append("\n");
				List<TopicConfiguration> topicConfigurations = getBucketNotificationConfigurationResponse.topicConfigurations();
				for (TopicConfiguration topicConfiguration : topicConfigurations) {
					bncs.append("Topic ARN = ");
					bncs.append(topicConfiguration.topicArn());
					bncs.append("\n");
					bncs.append("Event = ");
					bncs.append(topicConfiguration.eventsAsStrings());
				}
				bncs.append("\n\n");
				bncs.append("Lambda Configurations : ");
				bncs.append("\n");
				List<LambdaFunctionConfiguration> lambdaFunctionConfigurations = getBucketNotificationConfigurationResponse.lambdaFunctionConfigurations();
				for (LambdaFunctionConfiguration lambdaFunctionConfiguration : lambdaFunctionConfigurations) {
					bncs.append("Lambda ARN = ");
					bncs.append(lambdaFunctionConfiguration.lambdaFunctionArn());
					bncs.append("\n");
					bncs.append("Event = ");
					bncs.append(lambdaFunctionConfiguration.eventsAsStrings());
				}

				this.xssfHelper.setCell(row, bncs.toString());
			} catch(S3Exception skip) {
				this.xssfHelper.setCell(row, "");
			} catch(Exception e) {
				this.xssfHelper.setCell(row, "");
				e.printStackTrace();
			}
			
			try {
				GetBucketWebsiteResponse getBucketWebsiteResponse = this.amazonClients.s3Client.getBucketWebsite(GetBucketWebsiteRequest.builder().bucket(bucket.name()).build());
				this.xssfHelper.setCell(row, getBucketWebsiteResponse == null ? "" : getBucketWebsiteResponse.errorDocument() == null ? "" : getBucketWebsiteResponse.errorDocument().key());
				this.xssfHelper.setCell(row, getBucketWebsiteResponse == null ? "" : getBucketWebsiteResponse.indexDocument() == null ? "" : getBucketWebsiteResponse.indexDocument().suffix());
				
				if(getBucketWebsiteResponse != null) {
					RedirectAllRequestsTo redirectAllRequestsTo = getBucketWebsiteResponse.redirectAllRequestsTo();
					this.xssfHelper.setCell(row, redirectAllRequestsTo == null ? "" : redirectAllRequestsTo.hostName());
					this.xssfHelper.setCell(row, "");
					this.xssfHelper.setCell(row, redirectAllRequestsTo == null ? "" : this.getEnumName(redirectAllRequestsTo.protocol()));
					this.xssfHelper.setCell(row, "");
					this.xssfHelper.setCell(row, "");
					
					StringBuffer rors = new StringBuffer();
					List<RoutingRule> routingRules = getBucketWebsiteResponse.routingRules();
					for(RoutingRule routingRule : routingRules) {
						
						if(rors.length() > 0) rors.append("\n\n");
						
						Condition routingRuleCondition = routingRule.condition();
						
						rors.append("HttpErrorCodeReturnedEquals=");
						rors.append(routingRuleCondition.httpErrorCodeReturnedEquals());
						rors.append("\nKeyPrefixEquals=");
						rors.append(routingRuleCondition.keyPrefixEquals());
						
						Redirect rredirectRule = routingRule.redirect();
						rors.append("\nRedirectRule");
						rors.append("\nHostName=");
						rors.append(rredirectRule.hostName());
						rors.append("\nHttpRedirectCode=");
						rors.append(rredirectRule.httpRedirectCode());
						rors.append("\nProtocol=");
						rors.append(rredirectRule.protocol());
						rors.append("\nReplaceKeyPrefixWidth=");
						rors.append(rredirectRule.replaceKeyPrefixWith());
						rors.append("\nReplaceKeyWidth=");
						rors.append(rredirectRule.replaceKeyWith());
					}
					this.xssfHelper.setCell(row, rors.toString());
				} else {
					this.xssfHelper.setCell(row, "");
					this.xssfHelper.setCell(row, "");
					this.xssfHelper.setCell(row, "");
					this.xssfHelper.setCell(row, "");
					this.xssfHelper.setCell(row, "");
					this.xssfHelper.setCell(row, "");
				}

			} catch(S3Exception skip) {
				this.xssfHelper.setCell(row, "");
				this.xssfHelper.setCell(row, "");
				this.xssfHelper.setCell(row, "");
				this.xssfHelper.setCell(row, "");
				this.xssfHelper.setCell(row, "");
				this.xssfHelper.setCell(row, "");
				this.xssfHelper.setCell(row, "");
				this.xssfHelper.setCell(row, "");
			} catch(Exception e) {
				this.xssfHelper.setCell(row, "");
				this.xssfHelper.setCell(row, "");
				this.xssfHelper.setCell(row, "");
				this.xssfHelper.setCell(row, "");
				this.xssfHelper.setCell(row, "");
				this.xssfHelper.setCell(row, "");
				this.xssfHelper.setCell(row, "");
				this.xssfHelper.setCell(row, "");
				e.printStackTrace();
			}
			
			try {
				GetBucketAclResponse getBucketAclResponse = this.amazonClients.s3Client.getBucketAcl(GetBucketAclRequest.builder().bucket(bucket.name()).build());
				this.xssfHelper.setCell(row, getBucketAclResponse.owner().displayName());
				
				StringBuffer grts = new StringBuffer();
				List<Grant> grants = getBucketAclResponse.grants();
				for(Grant grant : grants) {
					Grantee grantee = grant.grantee();
					
					if(grts.length() > 0) grts.append("\n\n");
					grts.append("Identifier=");
					grts.append(grantee.id());
					grts.append("\nTypeIdentifier=");
					grts.append(grantee.typeAsString());
					grts.append("\nPermission=");
					grts.append(grant.permissionAsString());
				}
				this.xssfHelper.setCell(row, grts.toString());
			} catch(S3Exception skip) {
				this.xssfHelper.setCell(row, "");
				this.xssfHelper.setCell(row, "");
			} catch(Exception e) {
				this.xssfHelper.setCell(row, "");
				this.xssfHelper.setCell(row, "");
				e.printStackTrace();
			}
			
			try {
				StringBuffer lcs = new StringBuffer();
				GetBucketLifecycleConfigurationResponse getBucketLifecycleConfigurationResponse = this.amazonClients.s3Client.getBucketLifecycleConfiguration(GetBucketLifecycleConfigurationRequest.builder().bucket(bucket.name()).build());
				
				if(getBucketLifecycleConfigurationResponse != null) {
					List<LifecycleRule> rules = getBucketLifecycleConfigurationResponse.rules();
					for(LifecycleRule rule : rules) {
						AbortIncompleteMultipartUpload abortIncompleteMultipartUpload = rule.abortIncompleteMultipartUpload();
						if(lcs.length() > 0) lcs.append("\n\n");

						if(abortIncompleteMultipartUpload != null) {
							lcs.append("AbortInCompleteMultipartUpload=");
							lcs.append(Integer.toString(abortIncompleteMultipartUpload.daysAfterInitiation()));
							lcs.append("\n");
						}

						if(rule.expiration() != null) {
							lcs.append("ExpirationDate=");
							lcs.append(rule.expiration().date().toString());
							lcs.append("\n");
							lcs.append("ExpirationInDays=");
							lcs.append(Integer.toString(rule.expiration().days()));
							lcs.append("\n");
						}

						lcs.append("ID=");
						lcs.append(rule.id());
						if(rule.noncurrentVersionExpiration() != null) {
							lcs.append("\n");
							lcs.append("NonCurrentVersionExpirationInDays=");
							lcs.append(Integer.toString(rule.noncurrentVersionExpiration().noncurrentDays()));
						}
						lcs.append("\nNoncurrentVersionTransition");
						List<NoncurrentVersionTransition> noncurrentVersionTransitions = rule.noncurrentVersionTransitions();
						for(NoncurrentVersionTransition noncurrentVersionTransition : noncurrentVersionTransitions) {
							lcs.append("\nDays=");
							lcs.append(Integer.toString(noncurrentVersionTransition.noncurrentDays()));
							lcs.append("\nStorageClass=");
							lcs.append(noncurrentVersionTransition.storageClassAsString());
						}
						
						lcs.append("\nStatus=");
						lcs.append(this.getEnumName(rule.status()));
						
						lcs.append("\nTransitions");
						List<Transition> transitions = rule.transitions();
						for(Transition transition : transitions) {
							if(transition.date() != null) {
								lcs.append("\nDate=");
								lcs.append(transition.date().toString());
							}
							lcs.append("\nDays=");
							lcs.append(Integer.toString(transition.days()));
							lcs.append("\nStorageClass=");
							lcs.append(transition.storageClassAsString());
						}
					}
				}
				this.xssfHelper.setRightThinCell(row, lcs.toString());
			} catch(S3Exception skip) {
				this.xssfHelper.setRightThinCell(row, "");
			} catch(Exception e) {
				this.xssfHelper.setRightThinCell(row, "");
				e.printStackTrace();
			}
		}
	}
	
	
	private void makeACM() {
		
		XSSFRow row = null;
		
		ListCertificatesResponse listCertificatesResponse = this.amazonClients.acmClient.listCertificates();
		List<CertificateSummary> certificateSummaries = listCertificatesResponse.certificateSummaryList();
		for (CertificateSummary certificateSummary : certificateSummaries) {
			
			DescribeCertificateResponse describeCertificateResponse = this.amazonClients.acmClient.describeCertificate(DescribeCertificateRequest.builder().certificateArn(certificateSummary.certificateArn()).build());
			CertificateDetail certificateDetail = describeCertificateResponse.certificate();
			
			row = this.xssfHelper.createRow(this.acmSheet, 1);			
			this.xssfHelper.setLeftThinCell(row, Integer.toString(row.getRowNum() - 1));
			this.xssfHelper.setCell(row, certificateDetail.certificateArn());
			this.xssfHelper.setCell(row, certificateDetail.certificateAuthorityArn());
			this.xssfHelper.setCell(row, certificateDetail.domainName());
			this.xssfHelper.setCell(row, this.getEnumName(certificateDetail.status()));
			this.xssfHelper.setCell(row, certificateDetail.subject());
			
			StringBuffer sans = new StringBuffer();
			List<String> subjectAlternativeNames = certificateDetail.subjectAlternativeNames();
			for(String subjectAlternativeName : subjectAlternativeNames) {
				if(sans.length() > 0) sans.append("\n");
				sans.append(subjectAlternativeName);
			}
			this.xssfHelper.setCell(row, sans.toString());
			this.xssfHelper.setCell(row, this.getEnumName(certificateDetail.type()));
			this.xssfHelper.setCell(row, this.getEnumName(certificateDetail.keyAlgorithm()));
			
			StringBuffer kus = new StringBuffer();
			List<KeyUsage> keyUsages = certificateDetail.keyUsages();
			for(KeyUsage keyUsage : keyUsages) {
				if(kus.length() > 0) kus.append("\n");
				kus.append(keyUsage.name());
			}
			
			this.xssfHelper.setCell(row, kus.toString());
			this.xssfHelper.setCell(row, certificateDetail.serial());
			this.xssfHelper.setCell(row, certificateDetail.signatureAlgorithm());
			this.xssfHelper.setCell(row, this.getEnumName(certificateDetail.options().certificateTransparencyLoggingPreference()));
			
			this.xssfHelper.setCell(row, this.getDomainValidation(certificateDetail.domainValidationOptions()));
			
			StringBuffer ekus = new StringBuffer();
			List<ExtendedKeyUsage> extendedKeyUsages = certificateDetail.extendedKeyUsages();
			for(ExtendedKeyUsage extendedKeyUsage : extendedKeyUsages) {
				if(ekus.length() > 0) ekus.append("\n");
				ekus.append(extendedKeyUsage.name());
				if(extendedKeyUsage.oid() != null) {
					ekus.append(" (");
					ekus.append(extendedKeyUsage.oid());
					ekus.append(")");
				}
			}
			this.xssfHelper.setCell(row, ekus.toString());
			this.xssfHelper.setCell(row, this.getEnumName(certificateDetail.failureReason()));
			this.xssfHelper.setCell(row, certificateDetail.issuer());
			this.xssfHelper.setCell(row, this.getEnumName(certificateDetail.renewalEligibility()));
			RenewalSummary renewalSummary = certificateDetail.renewalSummary();
			this.xssfHelper.setCell(row, renewalSummary == null ? "" : this.getEnumName(renewalSummary.renewalStatus()));
			this.xssfHelper.setCell(row, renewalSummary == null ? "" : this.getDomainValidation(renewalSummary.domainValidationOptions()));
			this.xssfHelper.setCell(row, certificateDetail.createdAt() == null ? "" : certificateDetail.createdAt().toString());
			this.xssfHelper.setCell(row, certificateDetail.importedAt() == null ? "" : certificateDetail.importedAt().toString());
			this.xssfHelper.setCell(row, certificateDetail.inUseBy() == null ? "" : certificateDetail.inUseBy().toString());
			this.xssfHelper.setCell(row, certificateDetail.issuedAt() == null ? "" : certificateDetail.issuedAt().toString());
			this.xssfHelper.setCell(row, certificateDetail.revokedAt() == null ? "" : certificateDetail.revokedAt().toString());
			this.xssfHelper.setCell(row, this.getEnumName(certificateDetail.revocationReason()));
			this.xssfHelper.setCell(row, certificateDetail.notAfter() == null ? "" : certificateDetail.notAfter().toString());
			this.xssfHelper.setCell(row, certificateDetail.notBefore() == null ? "" : certificateDetail.notBefore().toString());
			
		}
	}
	
	private String getDomainValidation(List<DomainValidation> domainValidations) {
	
		StringBuffer dvs = new StringBuffer();
		for(DomainValidation domainValidation : domainValidations) {
			if(dvs.length() > 0) dvs.append("\n\n");
			dvs.append("DomainName=");
			dvs.append(domainValidation.domainName());
			dvs.append("\nResourceRecord=");
			dvs.append(domainValidation.resourceRecord());
			dvs.append("\nValidation Domain=");
			dvs.append(domainValidation.validationDomain());
			dvs.append("\nValidation Method=");
			dvs.append(domainValidation.validationMethod());

			List<String> validationEmails = domainValidation.validationEmails();
			if(validationEmails != null) {
				dvs.append("\nValidationEmail");
				for(String validationEmail : validationEmails) {
					dvs.append("\n\t-");
					dvs.append(validationEmail);
				}
			}
			dvs.append("\nValidationStatus=");
			dvs.append(domainValidation.validationStatus());
		}
		
		return dvs.toString();
	}
	
	private void makeKMS() {
		
		XSSFRow row = null;
		
		Map<String, AliasListEntry> aliasMap = new HashMap<>();
		ListAliasesResponse listAliasesResult = this.amazonClients.kmsClient.listAliases();
		List<AliasListEntry> aliasListEntries = listAliasesResult.aliases();
		for (AliasListEntry aliasListEntry : aliasListEntries) {
			
			if(aliasListEntry.targetKeyId() == null) {
				row = this.xssfHelper.createRow(this.kmsSheet, 1);			
				this.xssfHelper.setLeftThinCell(row, Integer.toString(row.getRowNum() - 1));
				this.xssfHelper.setCell(row, "");
				this.xssfHelper.setCell(row, aliasListEntry.aliasArn());
				this.xssfHelper.setCell(row, aliasListEntry.aliasName());
				this.xssfHelper.setCell(row, "");
				this.xssfHelper.setCell(row, "");
				this.xssfHelper.setCell(row, "");
				this.xssfHelper.setCell(row, "");
				this.xssfHelper.setCell(row, "");
				this.xssfHelper.setCell(row, "");
				this.xssfHelper.setCell(row, "");
				this.xssfHelper.setCell(row, "");
				this.xssfHelper.setCell(row, "");
				this.xssfHelper.setCell(row, "");
				this.xssfHelper.setCell(row, "");
				this.xssfHelper.setRightThinCell(row, "");
			} else {
				aliasMap.put(aliasListEntry.targetKeyId(), aliasListEntry);
			}
			
		}
		
		ListKeysResponse listKeysResult = this.amazonClients.kmsClient.listKeys();
		List<KeyListEntry> keyListEntries = listKeysResult.keys();
		for (KeyListEntry keyListEntry : keyListEntries) {
			
			row = this.xssfHelper.createRow(this.kmsSheet, 1);
			
			AliasListEntry aliasListEntry = aliasMap.get(keyListEntry.keyId());
			
			DescribeKeyResponse describeKeyResult = this.amazonClients.kmsClient.describeKey(DescribeKeyRequest.builder().keyId(keyListEntry.keyId()).build());
			KeyMetadata keyMetadata = describeKeyResult.keyMetadata();
			
			this.xssfHelper.setLeftThinCell(row, Integer.toString(row.getRowNum() - 1));
			this.xssfHelper.setCell(row, this.getEnumName(keyMetadata.keyManager()));
			this.xssfHelper.setCell(row, aliasListEntry == null ? "" : aliasListEntry.aliasArn());
			this.xssfHelper.setCell(row, aliasListEntry == null ? "" : aliasListEntry.aliasName());

			this.xssfHelper.setCell(row, keyMetadata.keyId());
			this.xssfHelper.setCell(row, keyMetadata.arn());
			this.xssfHelper.setCell(row, keyMetadata.awsAccountId());
			this.xssfHelper.setCell(row, keyMetadata.creationDate() == null ? "" : formatDate.format(keyMetadata.creationDate()));
			this.xssfHelper.setCell(row, keyMetadata.deletionDate() == null ? "" : formatDate.format(keyMetadata.deletionDate()));
			this.xssfHelper.setCell(row, keyMetadata.description());
			this.xssfHelper.setCell(row, keyMetadata.enabled().toString());
			this.xssfHelper.setCell(row, this.getEnumName(keyMetadata.expirationModel()));
			this.xssfHelper.setCell(row, this.getEnumName(keyMetadata.keyState()));
			this.xssfHelper.setCell(row, this.getEnumName(keyMetadata.keyUsage()));
			this.xssfHelper.setCell(row, this.getEnumName(keyMetadata.origin()));
			this.xssfHelper.setRightThinCell(row, keyMetadata.validTo() == null ? "" : keyMetadata.validTo().toString());
				
		}
		
	}
	
	private void makeElasticCache() {
		XSSFRow row = null;
    
		List<ReplicationGroup> replicationGroups = this.amazonClients.elastiCacheClient.describeReplicationGroups().replicationGroups();
		for (ReplicationGroup replicationGroup : replicationGroups) {
			int firstRowNum = 0;
			row = null;
			List<NodeGroup> nodeGroups = replicationGroup.nodeGroups();
			for (NodeGroup nodeGroup : nodeGroups) {
        
				List<NodeGroupMember> nodeGroupMembers = nodeGroup.nodeGroupMembers();
				for (NodeGroupMember nodeGroupMember : nodeGroupMembers) {
					CacheCluster cacheCluster = (CacheCluster)this.amazonClients.elastiCacheClient.describeCacheClusters(DescribeCacheClustersRequest.builder().cacheClusterId(nodeGroupMember.cacheClusterId()).build()).cacheClusters().get(0);
          
					row = this.xssfHelper.createRow(this.elastiCacheSheet, 1);
					if (firstRowNum == 0) {
						firstRowNum = row.getRowNum();
					}
					this.xssfHelper.setLeftThinCell(row, Integer.toString(row.getRowNum() - 2));
					this.xssfHelper.setCell(row, this.region.id());
					this.xssfHelper.setCell(row, replicationGroup.replicationGroupId());
					this.xssfHelper.setCell(row, cacheCluster.engine());
					this.xssfHelper.setCell(row, cacheCluster.engineVersion());
					this.xssfHelper.setCell(row, nodeGroup.primaryEndpoint() ==null ? "" : nodeGroup.primaryEndpoint().address() + ":" + nodeGroup.primaryEndpoint().port());
					this.xssfHelper.setCell(row, cacheCluster.cacheParameterGroup().cacheParameterGroupName() + "(" + cacheCluster.cacheParameterGroup().parameterApplyStatus() + ")");
					this.xssfHelper.setCell(row, cacheCluster.cacheSubnetGroupName());
          
					StringBuffer securityGroups = new StringBuffer();
					List<SecurityGroupMembership> securityGroupMemberships = cacheCluster.securityGroups();
					for (SecurityGroupMembership securityGroupMembership : securityGroupMemberships) {
						if (securityGroups.length() > 0) {
							securityGroups.append("\n");
						}
						securityGroups.append(securityGroupMembership.securityGroupId());
						securityGroups.append(" (");
						securityGroups.append(securityGroupMembership.status());
						securityGroups.append(")");
					}
					this.xssfHelper.setCell(row, securityGroups.toString());
					this.xssfHelper.setCell(row, cacheCluster.preferredMaintenanceWindow());
          
					this.xssfHelper.setCell(row, cacheCluster.cacheClusterId());
					this.xssfHelper.setCell(row, nodeGroupMember.preferredAvailabilityZone());
					this.xssfHelper.setCell(row, nodeGroupMember.currentRole());
					this.xssfHelper.setCell(row, cacheCluster.cacheNodeType());
					this.xssfHelper.setRightThinCell(row, nodeGroupMember.readEndpoint() == null ? "" : nodeGroupMember.readEndpoint().address() + ":" + nodeGroupMember.readEndpoint().port());
				}
			}
      
			if (row != null) {
				if (firstRowNum < row.getRowNum()) {
					this.elastiCacheSheet.addMergedRegion(new CellRangeAddress(firstRowNum, row.getRowNum(), 1, 1));
					this.elastiCacheSheet.addMergedRegion(new CellRangeAddress(firstRowNum, row.getRowNum(), 2, 2));
					this.elastiCacheSheet.addMergedRegion(new CellRangeAddress(firstRowNum, row.getRowNum(), 3, 3));
					this.elastiCacheSheet.addMergedRegion(new CellRangeAddress(firstRowNum, row.getRowNum(), 4, 4));
					this.elastiCacheSheet.addMergedRegion(new CellRangeAddress(firstRowNum, row.getRowNum(), 5, 5));
					this.elastiCacheSheet.addMergedRegion(new CellRangeAddress(firstRowNum, row.getRowNum(), 6, 6));
					this.elastiCacheSheet.addMergedRegion(new CellRangeAddress(firstRowNum, row.getRowNum(), 7, 7));
					this.elastiCacheSheet.addMergedRegion(new CellRangeAddress(firstRowNum, row.getRowNum(), 8, 8));
					this.elastiCacheSheet.addMergedRegion(new CellRangeAddress(firstRowNum, row.getRowNum(), 9, 9));
				}
				this.xssfHelper.setSubLastRowStyle(row);
			}
		}
	}
	
  
	private void makeRDSCluster() {
		
		XSSFRow row = null;
		
		DescribeDbClustersResponse describeDBClustersResult = this.amazonClients.rdsClient.describeDBClusters();
		List<DBCluster> dbClusters = describeDBClustersResult.dbClusters();
		for (DBCluster dbCluster : dbClusters) {

			row = this.xssfHelper.createRow(this.rdsClusterSheet, 1);
			this.xssfHelper.setLeftThinCell(row, Integer.toString(row.getRowNum() - 1));
			this.xssfHelper.setCell(row, "Cluster");
			this.xssfHelper.setCell(row, dbCluster.dbClusterIdentifier());
			
			StringBuffer avzs = new StringBuffer();
			List<String> availabilityZones = dbCluster.availabilityZones();
			for (String availabilityZone : availabilityZones) {
				if(avzs.length() > 0) avzs.append("\n");
				avzs.append(availabilityZone);
			}
			this.xssfHelper.setCell(row, avzs.toString());
			this.xssfHelper.setCell(row, dbCluster.dbSubnetGroup());
			
			StringBuffer dbclopgs = new StringBuffer();
			List<DBClusterOptionGroupStatus> dbClusterOptionGroupStatList = dbCluster.dbClusterOptionGroupMemberships();
			for(DBClusterOptionGroupStatus dbClusterOptionGroupStatus : dbClusterOptionGroupStatList) {
				if(dbclopgs.length() > 0) dbclopgs.append("\n");
				dbclopgs.append(dbClusterOptionGroupStatus.dbClusterOptionGroupName());
				dbclopgs.append(" (");
				dbclopgs.append(dbClusterOptionGroupStatus.status());
				dbclopgs.append(")");
			}
			this.xssfHelper.setCell(row, dbclopgs.toString());
			this.xssfHelper.setCell(row, dbCluster.dbClusterParameterGroup());
			this.xssfHelper.setCell(row, dbCluster.dbClusterResourceId());
			
			this.xssfHelper.setCell(row, dbCluster.databaseName());
			this.xssfHelper.setCell(row, dbCluster.characterSetName());

			this.xssfHelper.setCell(row, dbCluster.allocatedStorage().toString());
			this.xssfHelper.setCell(row, dbCluster.cloneGroupId());
			
			this.xssfHelper.setCell(row, dbCluster.backtrackWindow() == null ? "" : dbCluster.backtrackWindow().toString());
			this.xssfHelper.setCell(row, dbCluster.backupRetentionPeriod() == null ? "" : dbCluster.backupRetentionPeriod().toString());
			this.xssfHelper.setCell(row, dbCluster.backtrackConsumedChangeRecords() == null ? "" : dbCluster.backtrackConsumedChangeRecords().toString());

			this.xssfHelper.setCell(row, dbCluster.dbClusterArn());
			
			StringBuffer dbclrs = new StringBuffer();
			List<DBClusterRole> dbClusterRoles = dbCluster.associatedRoles();
			for(DBClusterRole dbClusterRole : dbClusterRoles) {
				if(dbclrs.length() > 0) dbclrs.append("\n");
				dbclrs.append("(");
				dbclrs.append(dbClusterRole.status());
				dbclrs.append(") ");
				dbclrs.append(dbClusterRole.roleArn());
			}
			this.xssfHelper.setCell(row, dbclrs.toString());
			
			this.xssfHelper.setCell(row, dbCluster.earliestBacktrackTime() == null ? "" : dbCluster.earliestBacktrackTime().toString());
			this.xssfHelper.setCell(row, dbCluster.earliestRestorableTime() == null ? "" : dbCluster.earliestRestorableTime().toString());
			
			StringBuffer ecle = new StringBuffer();
			List<String> enabledCloudwatchLogsExports = dbCluster.enabledCloudwatchLogsExports();
			for (String enabledCloudwatchLogsExport : enabledCloudwatchLogsExports) {
				if(ecle.length() > 0) ecle.append("\n");
				ecle.append(enabledCloudwatchLogsExport);
			}
			this.xssfHelper.setCell(row, ecle.toString());
			this.xssfHelper.setCell(row, dbCluster.endpoint());
			this.xssfHelper.setCell(row, dbCluster.engine());
			this.xssfHelper.setCell(row, dbCluster.engineVersion());
			
			this.xssfHelper.setCell(row, dbCluster.hostedZoneId());
			this.xssfHelper.setCell(row, dbCluster.iamDatabaseAuthenticationEnabled().toString());
			this.xssfHelper.setCell(row, dbCluster.kmsKeyId());
			this.xssfHelper.setCell(row, dbCluster.latestRestorableTime().toString());
			this.xssfHelper.setCell(row, dbCluster.masterUsername());
			this.xssfHelper.setCell(row, dbCluster.multiAZ().toString());
			this.xssfHelper.setCell(row, dbCluster.percentProgress());
			this.xssfHelper.setCell(row, dbCluster.port().toString());
			this.xssfHelper.setCell(row, dbCluster.preferredBackupWindow());
			this.xssfHelper.setCell(row, dbCluster.preferredMaintenanceWindow());
			this.xssfHelper.setCell(row, dbCluster.readerEndpoint());
			
			StringBuffer rrif = new StringBuffer();
			List<String> readReplicaIdentifiers = dbCluster.readReplicaIdentifiers();
			for(String readReplicaIdentifier : readReplicaIdentifiers) {
				if(rrif.length() > 0) rrif.append("\n");
				rrif.append(readReplicaIdentifier);
			}
			this.xssfHelper.setCell(row, rrif.toString());
			this.xssfHelper.setCell(row, dbCluster.replicationSourceIdentifier());
			this.xssfHelper.setCell(row, dbCluster.status());
			this.xssfHelper.setCell(row, dbCluster.storageEncrypted().toString());
			
			StringBuffer vsgs = new StringBuffer();
			List<VpcSecurityGroupMembership> vpcSecurityGroupMemberships = dbCluster.vpcSecurityGroups();
			for(VpcSecurityGroupMembership vpcSecurityGroupMembership : vpcSecurityGroupMemberships) {
				if(vsgs.length() > 0) vsgs.append("\n");
				vsgs.append(vpcSecurityGroupMembership.vpcSecurityGroupId());
				vsgs.append(" (");
				vsgs.append(vpcSecurityGroupMembership.status());
				vsgs.append(")");
			}
			this.xssfHelper.setCell(row, vsgs.toString());
			
			StringBuffer dcms = new StringBuffer();
			List<DBClusterMember> dbClusterMembers = dbCluster.dbClusterMembers();
			for (DBClusterMember dbClusterMember : dbClusterMembers) {
				if(dcms.length() > 0) dcms.append("\n");
				dcms.append("DBInstanceidentifier=");
				dcms.append(dbClusterMember.dbInstanceIdentifier());
				dcms.append(", IsClusterWriter=");
				dcms.append(dbClusterMember.isClusterWriter().toString());
				dcms.append(", DBClusterParameterGroupStatus=");
				dcms.append(dbClusterMember.dbClusterParameterGroupStatus());
				dcms.append(", PromotionTier");
				dcms.append(dbClusterMember.promotionTier().toString());
			}
			this.xssfHelper.setLeftRightThinCell(row, dcms.toString());

		}
		
	}
	
	
	private void makeRDSInstance() {
		
		DescribeDbInstancesResponse describeDBInstancesResult = this.amazonClients.rdsClient.describeDBInstances();
		List<DBInstance> dbInstances = describeDBInstancesResult.dbInstances();
		for (DBInstance dbInstance : dbInstances) {
			this.makeRDSDBInstances(dbInstance);
		}
		
	}
	
	private void makeRDSDBInstances(DBInstance dbInstance) {
		
		XSSFRow row = null;
		
		row = this.xssfHelper.createRow(this.rdsInstanceSheet, 1);
		this.xssfHelper.setLeftThinCell(row, Integer.toString(row.getRowNum() - 1));
		this.xssfHelper.setCell(row, dbInstance.dbClusterIdentifier() == null ? "Instance" : "Cluster Member");
		this.xssfHelper.setCell(row, dbInstance.dbClusterIdentifier());
		this.xssfHelper.setCell(row, dbInstance.dbInstanceIdentifier());
		this.xssfHelper.setCell(row, dbInstance.dbInstanceClass());
		this.xssfHelper.setCell(row, dbInstance.dbInstancePort().toString());
		this.xssfHelper.setCell(row, dbInstance.dbInstanceStatus());
		this.xssfHelper.setCell(row, dbInstance.availabilityZone());
		this.xssfHelper.setCell(row, dbInstance.dbSubnetGroup().dbSubnetGroupName());
		
		StringBuffer dbsgms = new StringBuffer();
		List<DBSecurityGroupMembership> dbSecurityGroups = dbInstance.dbSecurityGroups();
		for (DBSecurityGroupMembership dbSecurityGroupMembership : dbSecurityGroups) {
			if(dbsgms.length() > 0) dbsgms.append("\n");
			dbsgms.append(dbSecurityGroupMembership.dbSecurityGroupName());
			dbsgms.append(" (");
			dbsgms.append(dbSecurityGroupMembership.status());
			dbsgms.append(")");
		}
		
		this.xssfHelper.setCell(row, dbsgms.toString());
		
		StringBuffer dbpgs = new StringBuffer();
		List<DBParameterGroupStatus> dbParameterGroupStatList = dbInstance.dbParameterGroups();
		for (DBParameterGroupStatus dbParameterGroupStatus : dbParameterGroupStatList) {
			if(dbpgs.length() > 0) dbpgs.append("\n");
			dbpgs.append(dbParameterGroupStatus.dbParameterGroupName());
			dbpgs.append(" (");
			dbpgs.append(dbParameterGroupStatus.parameterApplyStatus());
			dbpgs.append(")");
		}
		
		this.xssfHelper.setCell(row, dbpgs.toString());
		this.xssfHelper.setCell(row, dbInstance.dbName());
		this.xssfHelper.setCell(row, dbInstance.characterSetName());

		this.xssfHelper.setCell(row, dbInstance.caCertificateIdentifier());
		
		this.xssfHelper.setCell(row, dbInstance.autoMinorVersionUpgrade().toString());
		this.xssfHelper.setCell(row, dbInstance.backupRetentionPeriod().toString());

		this.xssfHelper.setCell(row, dbInstance.copyTagsToSnapshot().toString());

		this.xssfHelper.setCell(row, dbInstance.dbiResourceId());
		this.xssfHelper.setCell(row, dbInstance.allocatedStorage().toString());
		this.xssfHelper.setCell(row, dbInstance.dbInstanceArn());
		
		StringBuffer dmss = new StringBuffer();
		List<DomainMembership> domainMemberships = dbInstance.domainMemberships();
		for (DomainMembership domainMembership : domainMemberships) {
			if(dmss.length() > 0) dmss.append("\n");
			dmss.append("Domain=");
			dmss.append(domainMembership.domain());
			dmss.append(", FQDN=");
			dmss.append(domainMembership.fqdn());
			dmss.append(", IAMRole=");
			dmss.append(domainMembership.iamRoleName());
			dmss.append(", Status=");
			dmss.append(domainMembership.status());
		}
		this.xssfHelper.setCell(row, dmss.toString());
		
		StringBuffer ecle = new StringBuffer();
		List<String> enabledCloudwatchLogsExports = dbInstance.enabledCloudwatchLogsExports();
		for (String enabledCloudwatchLogsExport : enabledCloudwatchLogsExports) {
			if(ecle.length() > 0) ecle.append("\n");
			ecle.append(enabledCloudwatchLogsExport);
		}
		this.xssfHelper.setCell(row, ecle.toString());
		
		Endpoint endpoint = dbInstance.endpoint();
		this.xssfHelper.setCell(row, endpoint.address());
		this.xssfHelper.setCell(row, endpoint.hostedZoneId());
		this.xssfHelper.setCell(row, endpoint.port().toString());
		this.xssfHelper.setCell(row, dbInstance.engine());
		this.xssfHelper.setCell(row, dbInstance.engineVersion());
		this.xssfHelper.setCell(row, dbInstance.enhancedMonitoringResourceArn());
		this.xssfHelper.setCell(row, dbInstance.iamDatabaseAuthenticationEnabled().toString());
		this.xssfHelper.setCell(row, dbInstance.instanceCreateTime().toString());
		this.xssfHelper.setCell(row, dbInstance.iops() == null ? "" : dbInstance.iops().toString());
		this.xssfHelper.setCell(row, dbInstance.kmsKeyId());
		this.xssfHelper.setCell(row, dbInstance.latestRestorableTime() == null ? "" : dbInstance.latestRestorableTime().toString());
		this.xssfHelper.setCell(row, dbInstance.licenseModel());
		this.xssfHelper.setCell(row, dbInstance.masterUsername());
		this.xssfHelper.setCell(row, dbInstance.monitoringInterval() == null ? "" : dbInstance.monitoringInterval().toString());
		this.xssfHelper.setCell(row, dbInstance.monitoringRoleArn());
		this.xssfHelper.setCell(row, dbInstance.multiAZ().toString());
		this.xssfHelper.setCell(row, dbInstance.optionGroupMemberships().toString());
		this.xssfHelper.setCell(row, dbInstance.performanceInsightsEnabled().toString());
		this.xssfHelper.setCell(row, dbInstance.performanceInsightsKMSKeyId());
		this.xssfHelper.setCell(row, dbInstance.performanceInsightsRetentionPeriod() == null ? "" : dbInstance.performanceInsightsRetentionPeriod().toString());
		this.xssfHelper.setCell(row, dbInstance.preferredBackupWindow());
		this.xssfHelper.setCell(row, dbInstance.preferredMaintenanceWindow());
		this.xssfHelper.setCell(row, dbInstance.processorFeatures().toString());
		this.xssfHelper.setCell(row, dbInstance.promotionTier() == null ? "" : dbInstance.promotionTier().toString());
		this.xssfHelper.setCell(row, dbInstance.publiclyAccessible().toString());
		
		StringBuffer rdcis = new StringBuffer();
		List<String> readReplicaDBClusterIdentifiers = dbInstance.readReplicaDBClusterIdentifiers();
		for(String readReplicaDBClusterIdentifier : readReplicaDBClusterIdentifiers) {
			if(rdcis.length() > 0) rdcis.append("\n");
			rdcis.append(readReplicaDBClusterIdentifier);
		}
		this.xssfHelper.setCell(row, rdcis.toString());
		
		StringBuffer rdiis = new StringBuffer();
		List<String> readReplicaDBInstanceIdentifiers = dbInstance.readReplicaDBInstanceIdentifiers();
		for(String readReplicaDBInstanceIdentifier : readReplicaDBInstanceIdentifiers) {
			if(rdiis.length() > 0) rdiis.append("\n");
			rdiis.append(readReplicaDBInstanceIdentifier);
		}
		
		this.xssfHelper.setCell(row, rdiis.toString());
		this.xssfHelper.setCell(row, dbInstance.readReplicaSourceDBInstanceIdentifier());
		this.xssfHelper.setCell(row, dbInstance.secondaryAvailabilityZone());
		this.xssfHelper.setCell(row, dbInstance.statusInfos().toString());
		this.xssfHelper.setCell(row, dbInstance.storageEncrypted().toString());
		this.xssfHelper.setCell(row, dbInstance.storageType());
		this.xssfHelper.setCell(row, dbInstance.tdeCredentialArn());
		this.xssfHelper.setCell(row, dbInstance.timezone());
		
		StringBuffer vsgs = new StringBuffer();
		List<VpcSecurityGroupMembership> vpcSecurityGroupMemberships = dbInstance.vpcSecurityGroups();
		for(VpcSecurityGroupMembership vpcSecurityGroupMembership : vpcSecurityGroupMemberships) {
			if(vsgs.length() > 0) vsgs.append("\n");
			vsgs.append(vpcSecurityGroupMembership.vpcSecurityGroupId());
			vsgs.append(" (");
			vsgs.append(vpcSecurityGroupMembership.status());
			vsgs.append(")");
		}
		this.xssfHelper.setRightThinCell(row, vsgs.toString());
	}
	
	
	private void makeEBS() {
		XSSFRow row = null;
    
		List<Volume> volumes = this.amazonClients.ec2Client.describeVolumes().volumes();
		for (Volume volume : volumes) {
			row = this.xssfHelper.createRow(this.ebsSheet, 1);
			this.xssfHelper.setLeftThinCell(row, Integer.toString(row.getRowNum() - 1));
			this.xssfHelper.setCell(row, this.region.id());
			this.xssfHelper.setCell(row, getNameTagValue(volume.tags()));
			this.xssfHelper.setCell(row, volume.volumeId());
			this.xssfHelper.setCell(row, this.getEnumName(volume.volumeType()));
			this.xssfHelper.setCell(row, volume.size().toString());
			this.xssfHelper.setCell(row, volume.iops() == null ? "-" : volume.iops().toString());
			this.xssfHelper.setCell(row, volume.availabilityZone());
			this.xssfHelper.setCell(row, volume.snapshotId());
      
			StringBuffer volumeAttachmentBuffer = new StringBuffer();
			List<VolumeAttachment> volumeAttachments = volume.attachments();
			for (VolumeAttachment volumeAttachment : volumeAttachments) {
				Instance instance = ((Reservation)this.amazonClients.ec2Client.describeInstances(DescribeInstancesRequest.builder().instanceIds(new String[] { volumeAttachment.instanceId() }).build()).reservations().get(0)).instances().get(0);
				if (volumeAttachmentBuffer.length() > 0) {
					volumeAttachmentBuffer.append("\n");
				}
				volumeAttachmentBuffer.append(volumeAttachment.instanceId());
				volumeAttachmentBuffer.append("(");
				volumeAttachmentBuffer.append(getNameTagValue(instance.tags()));
				volumeAttachmentBuffer.append("):");
				volumeAttachmentBuffer.append(volumeAttachment.device());
				volumeAttachmentBuffer.append("(");
				volumeAttachmentBuffer.append(volumeAttachment.state());
				volumeAttachmentBuffer.append(")");
			}
			this.xssfHelper.setCell(row, volumeAttachmentBuffer.toString());
      
			this.xssfHelper.setCell(row, volume.encrypted().booleanValue() ? "Ecrypted" : "Not Encrypted");
			this.xssfHelper.setCell(row, volume.kmsKeyId());
			this.xssfHelper.setRightThinCell(row, this.getAllTagValue(volume.tags()));
		}
	}
  
	
	private void makeClassicELB() {
		XSSFRow row = null;
    
		List<LoadBalancerDescription> loadBalancerDescriptions = this.amazonClients.elasticLoadBalancingClient.describeLoadBalancers().loadBalancerDescriptions();
		for (LoadBalancerDescription loadBalancerDescription : loadBalancerDescriptions) {
			row = this.xssfHelper.createRow(this.classicElbSheet, 1);
			int startRowNum = row.getRowNum();
			this.xssfHelper.setLeftThinCell(row, Integer.toString(row.getRowNum() - 1));
			this.xssfHelper.setCell(row, this.region.id());
			this.xssfHelper.setCell(row, loadBalancerDescription.vpcId());
			this.xssfHelper.setCell(row, "");
			StringBuffer avzBuffer = new StringBuffer();
			List<String> subnetIds = loadBalancerDescription.subnets();
			for (String subnetId : subnetIds) {
				Subnet subnet = this.amazonClients.ec2Client.describeSubnets(DescribeSubnetsRequest.builder().subnetIds(new String[] { subnetId }).build()).subnets().get(0);
				if (avzBuffer.length() > 0) {
					avzBuffer.append("\n");
				}
				avzBuffer.append(subnet.availabilityZone());
				avzBuffer.append("(");
				avzBuffer.append(subnet.subnetId());
				avzBuffer.append(")");
			}
			this.xssfHelper.setCell(row, avzBuffer.toString());
			this.xssfHelper.setCell(row, loadBalancerDescription.loadBalancerName());
			this.xssfHelper.setCell(row, loadBalancerDescription.dnsName());
      
			StringBuffer instanceBuffer = new StringBuffer();
			List<software.amazon.awssdk.services.elasticloadbalancing.model.Instance> instances = loadBalancerDescription.instances();
			for (software.amazon.awssdk.services.elasticloadbalancing.model.Instance instance : instances) {
				Reservation reservation = (Reservation)this.amazonClients.ec2Client.describeInstances(DescribeInstancesRequest.builder().instanceIds(new String[] { instance.instanceId() }).build()).reservations().get(0);
				Instance ec2Instance = reservation.instances().get(0);
				if (instanceBuffer.length() > 0) {
					instanceBuffer.append("\n");
				}
				instanceBuffer.append(ec2Instance.instanceId() + " [" + getNameTagValue(ec2Instance.tags()) + "]");
			}

			this.xssfHelper.setCell(row, instanceBuffer.toString());
			this.xssfHelper.setSubHeadLeftThinCell(row, "Load Balancer Protocol");
			this.xssfHelper.setSubHeadCell(row, "Load Balancer Port");
			this.xssfHelper.setSubHeadCell(row, "Instance Protocol");
			this.xssfHelper.setSubHeadRightThinCell(row, "Instance Port");
			this.xssfHelper.setRightThinCell(row, loadBalancerDescription.healthCheck().target());
      
			List<ListenerDescription> listenerDescriptions = loadBalancerDescription.listenerDescriptions();
			for (ListenerDescription listenerDescription : listenerDescriptions) {
				Listener listener = listenerDescription.listener();
        
				row = this.xssfHelper.createRow(this.classicElbSheet, 1);
				this.xssfHelper.setLeftThinCell(row, Integer.toString(row.getRowNum() - 1));
				this.xssfHelper.setCell(row, "");
				this.xssfHelper.setCell(row, "");
				this.xssfHelper.setCell(row, "");
				this.xssfHelper.setCell(row, "");
				this.xssfHelper.setCell(row, "");
				this.xssfHelper.setCell(row, "");
				this.xssfHelper.setCell(row, "");
				this.xssfHelper.setLeftThinCell(row, listener.protocol());
				this.xssfHelper.setCell(row, listener.loadBalancerPort().toString());
				this.xssfHelper.setCell(row, listener.instanceProtocol());
				this.xssfHelper.setRightThinCell(row, listener.instancePort().toString());
				this.xssfHelper.setRightThinCell(row, "");
			}
			if (startRowNum < row.getRowNum()) {
				this.classicElbSheet.addMergedRegionUnsafe(new CellRangeAddress(startRowNum, row.getRowNum(), 1, 1));
				this.classicElbSheet.addMergedRegionUnsafe(new CellRangeAddress(startRowNum, row.getRowNum(), 2, 2));
				this.classicElbSheet.addMergedRegionUnsafe(new CellRangeAddress(startRowNum, row.getRowNum(), 3, 3));
				this.classicElbSheet.addMergedRegionUnsafe(new CellRangeAddress(startRowNum, row.getRowNum(), 4, 4));
				this.classicElbSheet.addMergedRegionUnsafe(new CellRangeAddress(startRowNum, row.getRowNum(), 5, 5));
				this.classicElbSheet.addMergedRegionUnsafe(new CellRangeAddress(startRowNum, row.getRowNum(), 6, 6));
				this.classicElbSheet.addMergedRegionUnsafe(new CellRangeAddress(startRowNum, row.getRowNum(), 7, 7));
				this.classicElbSheet.addMergedRegionUnsafe(new CellRangeAddress(startRowNum, row.getRowNum(), 12, 12));
			}
			this.xssfHelper.setSubLastRowStyle(row);
		}
	}
  
	
	private void makeOtherELB() {
		XSSFRow row = null;
	  	  
		DescribeLoadBalancersResponse describeLoadBalancersResult = this.amazonClients.elasticLoadBalancingV2Client.describeLoadBalancers();
		List<LoadBalancer> loadBalancers = describeLoadBalancersResult.loadBalancers();
		for (LoadBalancer loadBalancer : loadBalancers) {
		  
			row = this.xssfHelper.createRow(this.otherElbSheet, 1);
			//int startRowNum = row.getRowNum();
			this.xssfHelper.setLeftThinCell(row, Integer.toString(row.getRowNum() - 1));
			this.xssfHelper.setCell(row, this.region.id());
			this.xssfHelper.setCell(row, loadBalancer.vpcId());
	      
			StringBuffer azs = new StringBuffer();
			List<AvailabilityZone> availabilityZones = loadBalancer.availabilityZones();
			for (AvailabilityZone availabilityZone : availabilityZones) {
				if(azs.length() > 0) azs.append("\n");
				azs.append("Zone Name = ");
				azs.append(availabilityZone.zoneName());
				azs.append("\n Subnet ID = ");
				azs.append(availabilityZone.subnetId());
			
				StringBuffer lba = new StringBuffer();
				List<LoadBalancerAddress> loadBalancerAddresses = availabilityZone.loadBalancerAddresses();
				if(loadBalancerAddresses != null) {
					for(LoadBalancerAddress loadBalancerAddress : loadBalancerAddresses) {
						lba.append("\nIp : ");
						lba.append(loadBalancerAddress.ipAddress());
					}
				}
				azs.append(lba.toString() + "");
			}
	      
			this.xssfHelper.setCell(row, azs.toString());
			this.xssfHelper.setCell(row, this.getEnumName(loadBalancer.type()));
			this.xssfHelper.setCell(row, this.getEnumName(loadBalancer.scheme()));
			this.xssfHelper.setCell(row, this.getEnumName(loadBalancer.ipAddressType()));
			this.xssfHelper.setCell(row, loadBalancer.loadBalancerName());
			this.xssfHelper.setCell(row, loadBalancer.canonicalHostedZoneId());
			String stsReason = loadBalancer.state().reason();
			this.xssfHelper.setCell(row, loadBalancer.state().code() + (stsReason == null ? "" : "(" + stsReason + ")"));
	      
			StringBuffer sgs = new StringBuffer();
			List<String> securityGroups = loadBalancer.securityGroups();
			if(securityGroups != null) {
				for(String securityGroup : securityGroups) {
					sgs.append("\n");
					sgs.append(securityGroup);
				}
			}

			this.xssfHelper.setCell(row, sgs.toString() + "");
			this.xssfHelper.setCell(row, loadBalancer.dnsName());
			this.xssfHelper.setCell(row, "");
			this.xssfHelper.setCell(row, "");
			this.xssfHelper.setCell(row, loadBalancer.loadBalancerArn());
	      
			this.otherElbSheet.addMergedRegion(new CellRangeAddress(row.getRowNum(), row.getRowNum(), 11, 13));
	      
			StringBuffer tsb = new StringBuffer();
	      
			DescribeTagsResponse describeTagsResult = this.amazonClients.elasticLoadBalancingV2Client.describeTags(DescribeTagsRequest.builder().resourceArns(loadBalancer.loadBalancerArn()).build());
			List<TagDescription> tagDescriptions = describeTagsResult.tagDescriptions();
			for(TagDescription tagDescription : tagDescriptions) {
				for(software.amazon.awssdk.services.elasticloadbalancingv2.model.Tag tag : tagDescription.tags()) {
					if(tsb.length() > 0) tsb.append("\n");
					tsb.append("Key=");
					tsb.append(tag.key());
					tsb.append(", Value=");
					tsb.append(tag.value());
				}  
			}
	      
			this.xssfHelper.setRightThinCell(row, tsb.toString() + "");
	      
			// Listener Start
			// Listener Header Row
			row = this.xssfHelper.createRow(this.otherElbSheet, 1);
			this.xssfHelper.setLeftThinCell(row, Integer.toString(row.getRowNum() - 1));
	    	this.xssfHelper.setBrownBoldCell(row, "Listener Protocol");
	    	this.xssfHelper.setBrownBoldCell(row, "Listener Port");
	    	this.xssfHelper.setBrownBoldCell(row, "SSL Policy");
	    	this.xssfHelper.setBrownBoldCell(row, "Certificate");
	    	this.xssfHelper.setBrownBoldCell(row, "");
	    	this.xssfHelper.setBrownBoldCell(row, "");
	    	this.xssfHelper.setBrownBoldCell(row, "");
	    	this.xssfHelper.setBrownBoldCell(row, "Listener ARN");
	    	this.xssfHelper.setBrownBoldCell(row, "");
	    	this.xssfHelper.setBrownBoldCell(row, "");
	    	this.xssfHelper.setBrownBoldCell(row, "");
	    	this.xssfHelper.setBrownBoldCell(row, "");
	    	this.xssfHelper.setBrownBoldCell(row, "");
	    	this.xssfHelper.setBrownBoldCell(row, "");
	    	this.xssfHelper.setBrownBoldRightThinCell(row, "");
	    	
	    	this.otherElbSheet.addMergedRegion(new CellRangeAddress(row.getRowNum(), row.getRowNum(), 4, 7));
	    	this.otherElbSheet.addMergedRegion(new CellRangeAddress(row.getRowNum(), row.getRowNum(), 8, 15));
	    	
	    	// Listener Data Row
	    	DescribeListenersResponse describeListenersResult = this.amazonClients.elasticLoadBalancingV2Client.describeListeners(DescribeListenersRequest.builder().loadBalancerArn(loadBalancer.loadBalancerArn()).build());
	      
	    	List<software.amazon.awssdk.services.elasticloadbalancingv2.model.Listener> listeners = describeListenersResult.listeners();
	    	for (software.amazon.awssdk.services.elasticloadbalancingv2.model.Listener listener : listeners) {
	    	
	    		row = this.xssfHelper.createRow(this.otherElbSheet, 1);
	    		this.xssfHelper.setLeftThinCell(row, Integer.toString(row.getRowNum() - 1));
	    		this.xssfHelper.setSubHeadCell(row, this.getEnumName(listener.protocol()));
	    		this.xssfHelper.setSubHeadCell(row, listener.port().toString());

	    		if("network".equals(this.getEnumName(loadBalancer.type()))) {
	    			List<Action> actions = listener.defaultActions();
	    			if(actions != null) {
	    				this.makeOtherELBTargetGroupActionHeader("Default");
	    				this.makeOtherELBActions(actions);
	    			}
	    		}
	    	
	    		if("application".equals(this.getEnumName(loadBalancer.type()))) {
	    			DescribeRulesResponse describeRulesResult = this.amazonClients.elasticLoadBalancingV2Client.describeRules(DescribeRulesRequest.builder().listenerArn(listener.listenerArn()).build());
	    			List<Rule> rules = describeRulesResult.rules();
	    			for (Rule rule : rules) {
	    				this.makeOtherELBRules(rule);
	    				this.makeOtherELBTargetGroupActionHeader("Rule");
	    				this.makeOtherELBActions(rule.actions());
	    			}
	    		}
	    	
	    		this.xssfHelper.setSubHeadCell(row, listener.sslPolicy());
	    	
	    		StringBuffer scer = new StringBuffer();
	    		List<Certificate> certificates = listener.certificates();
	    		if(certificates != null) {
	    			for(Certificate certificate : certificates) {
	    				if(scer.length() > 0) scer.append("\n");
	    				scer.append("Certificate ARN=");
	    				scer.append(certificate.certificateArn());
	    				scer.append(", isDefault=");
	    				scer.append(certificate.isDefault());
	    			}
	    		}
	    		this.xssfHelper.setSubHeadCell(row, scer.toString() + "");
	    		this.xssfHelper.setSubHeadCell(row, "");
	    		this.xssfHelper.setSubHeadCell(row, "");
	    		this.xssfHelper.setSubHeadCell(row, "");
	    		this.xssfHelper.setSubHeadCell(row, listener.listenerArn());
	    		this.xssfHelper.setSubHeadCell(row, "");
	    		this.xssfHelper.setSubHeadCell(row, "");
	    		this.xssfHelper.setSubHeadCell(row, "");
	    		this.xssfHelper.setSubHeadCell(row, "");
	    		this.xssfHelper.setSubHeadCell(row, "");
	    		this.xssfHelper.setSubHeadCell(row, "");
	    		this.xssfHelper.setSubHeadRightThinCell(row, "");
	    	
	    		this.otherElbSheet.addMergedRegion(new CellRangeAddress(row.getRowNum(), row.getRowNum(), 4, 7));
	    		this.otherElbSheet.addMergedRegion(new CellRangeAddress(row.getRowNum(), row.getRowNum(), 8, 15));
	    	
	    	}
	    	this.xssfHelper.setSubLastRowStyle(this.otherElbSheet.getRow(this.otherElbSheet.getLastRowNum()));
		}
	}
  
	private void makeOtherELBRules(Rule rule) {
		XSSFRow row = null;
	  
		row = this.xssfHelper.createRow(this.otherElbSheet, 1);
		this.xssfHelper.setLeftThinCell(row, Integer.toString(row.getRowNum() - 1));
		this.xssfHelper.setSecondSubHeadCell(row, "is Default Rule");
		this.xssfHelper.setSecondSubHeadCell(row, "Rule Priority");
		this.xssfHelper.setSecondSubHeadCell(row, "Rule Condition");
		this.xssfHelper.setSecondSubHeadCell(row, "");
		this.xssfHelper.setSecondSubHeadCell(row, "");
		this.xssfHelper.setSecondSubHeadCell(row, "Rule ARN");
		this.xssfHelper.setSecondSubHeadCell(row, "");
		this.xssfHelper.setSecondSubHeadCell(row, "");
		this.xssfHelper.setSecondSubHeadCell(row, "");
		this.xssfHelper.setSecondSubHeadCell(row, "");
		this.xssfHelper.setSecondSubHeadCell(row, "");
		this.xssfHelper.setSecondSubHeadCell(row, "");
		this.xssfHelper.setSecondSubHeadCell(row, "");
		this.xssfHelper.setSecondSubHeadCell(row, "");
		this.xssfHelper.setSecondSubHeadRightThinCell(row, "");

		this.otherElbSheet.addMergedRegion(new CellRangeAddress(row.getRowNum(), row.getRowNum(), 3, 5));
		this.otherElbSheet.addMergedRegion(new CellRangeAddress(row.getRowNum(), row.getRowNum(), 6, 15));
      
		StringBuffer ruleCs = new StringBuffer();
		List<RuleCondition> ruleConditions = rule.conditions();
		for (RuleCondition ruleCondition : ruleConditions) {
		  
			ruleCs.append(ruleCondition.field());
			ruleCs.append("=");
		  
			StringBuffer fieldVs = new StringBuffer();
			List<String> fieldValues = ruleCondition.values();
			for(String fieldValue : fieldValues) {
				if(fieldVs.length() > 0) fieldVs.append(", ");
				fieldVs.append(fieldValue);  
			}
	      
			ruleCs.append(fieldVs.toString());
		}
	  
		row = this.xssfHelper.createRow(this.otherElbSheet, 1);
		this.xssfHelper.setLeftThinCell(row, Integer.toString(row.getRowNum() - 1));
		this.xssfHelper.setCell(row, rule.isDefault().toString());
		this.xssfHelper.setCell(row, rule.priority());
		this.xssfHelper.setCell(row, ruleCs.toString());
		this.xssfHelper.setCell(row, "");
		this.xssfHelper.setCell(row, "");
		this.xssfHelper.setCell(row, rule.ruleArn());
		this.xssfHelper.setCell(row, "");
		this.xssfHelper.setCell(row, "");
		this.xssfHelper.setCell(row, "");
		this.xssfHelper.setCell(row, "");
		this.xssfHelper.setCell(row, "");
		this.xssfHelper.setCell(row, "");
		this.xssfHelper.setCell(row, "");
		this.xssfHelper.setCell(row, "");
		this.xssfHelper.setRightThinCell(row, "");
      
		this.otherElbSheet.addMergedRegion(new CellRangeAddress(row.getRowNum(), row.getRowNum(), 3, 5));
		this.otherElbSheet.addMergedRegion(new CellRangeAddress(row.getRowNum(), row.getRowNum(), 6, 15));

	}
  
	private void makeOtherELBActions(List<Action> actions) {
		
	    for(Action action : actions) {	    		    
		    // Listener Target Group Start
	    	DescribeTargetGroupsResponse describeTargetGroupsResult = this.amazonClients.elasticLoadBalancingV2Client.describeTargetGroups(DescribeTargetGroupsRequest.builder().targetGroupArns(action.targetGroupArn()).build());
	    	List<TargetGroup> targetGroups = describeTargetGroupsResult.targetGroups();
	    	for(TargetGroup targetGroup : targetGroups) {

  	    	this.makeOtherELBTargetGroupHeader();
  	    	
	    		StringBuffer tga = new StringBuffer();
	    		DescribeTargetGroupAttributesResponse describeTargetGroupAttributesResult = this.amazonClients.elasticLoadBalancingV2Client.describeTargetGroupAttributes(DescribeTargetGroupAttributesRequest.builder().targetGroupArn(targetGroup.targetGroupArn()).build());
	    		List<TargetGroupAttribute> targetGroupAttributes = describeTargetGroupAttributesResult.attributes();
	    		for(TargetGroupAttribute targetGroupAttribute : targetGroupAttributes) {
	    			if(tga.length() > 0) tga.append("\n");
	    			tga.append(targetGroupAttribute.key());
	    			tga.append("=");
	    			tga.append(targetGroupAttribute.value());
	    		}
	    		
	    		this.makeOtherELBTargetGroup(action, targetGroup, tga.toString());	    	    		
	    		
	    		this.makeOtherELBTargetHealthHeader();
	    		DescribeTargetHealthResponse describeTargetHealthResult = this.amazonClients.elasticLoadBalancingV2Client.describeTargetHealth(DescribeTargetHealthRequest.builder().targetGroupArn(targetGroup.targetGroupArn()).build());
	    		List<TargetHealthDescription> targetHealthDescriptions = describeTargetHealthResult.targetHealthDescriptions();
	    		for(TargetHealthDescription targetHealthDescription : targetHealthDescriptions) {
	    			this.makeOtherELBTargetHealth(targetHealthDescription);
	    		}
	    		
	    	}
	    }
	}
  
	private void makeOtherELBTargetGroupActionHeader(String actionName) {
		XSSFRow row = this.xssfHelper.createRow(this.otherElbSheet, 1);
		this.xssfHelper.setLeftThinCell(row, Integer.toString(row.getRowNum() - 1));
		this.xssfHelper.setSecondSubHeadCell(row, actionName + " Actions");
		this.xssfHelper.setSecondSubHeadCell(row, "");
		this.xssfHelper.setSecondSubHeadCell(row, "");
		this.xssfHelper.setSecondSubHeadCell(row, "");
		this.xssfHelper.setSecondSubHeadCell(row, "");
		this.xssfHelper.setSecondSubHeadCell(row, "");
		this.xssfHelper.setSecondSubHeadCell(row, "");
		this.xssfHelper.setSecondSubHeadCell(row, "");
		this.xssfHelper.setSecondSubHeadCell(row, "");
		this.xssfHelper.setSecondSubHeadCell(row, "");
		this.xssfHelper.setSecondSubHeadCell(row, "");
		this.xssfHelper.setSecondSubHeadCell(row, "");
		this.xssfHelper.setSecondSubHeadCell(row, "");
		this.xssfHelper.setSecondSubHeadCell(row, "");
		this.xssfHelper.setSecondSubHeadRightThinCell(row, "");
		this.otherElbSheet.addMergedRegion(new CellRangeAddress(row.getRowNum(), row.getRowNum(), 1, 15));
	}
  
	private void makeOtherELBTargetHealthHeader() {
		XSSFRow row = this.xssfHelper.createRow(this.otherElbSheet, 1);
		this.xssfHelper.setLeftThinCell(row, Integer.toString(row.getRowNum() - 1));
		this.xssfHelper.setSecondSubHeadCell(row, "Availability Zone");
		this.xssfHelper.setSecondSubHeadCell(row, "Target Id");
		this.xssfHelper.setSecondSubHeadCell(row, "Traffic Port");
		this.xssfHelper.setSecondSubHeadCell(row, "HealthCheck Port");
		this.xssfHelper.setSecondSubHeadCell(row, "Health State");
		this.xssfHelper.setSecondSubHeadCell(row, "Health Reason");
		this.xssfHelper.setSecondSubHeadCell(row, "Health Description");
		this.xssfHelper.setSecondSubHeadCell(row, "Remark");
		this.xssfHelper.setSecondSubHeadCell(row, "");
		this.xssfHelper.setSecondSubHeadCell(row, "");
		this.xssfHelper.setSecondSubHeadCell(row, "");
		this.xssfHelper.setSecondSubHeadCell(row, "");
		this.xssfHelper.setSecondSubHeadCell(row, "");
		this.xssfHelper.setSecondSubHeadCell(row, "");
		this.xssfHelper.setSecondSubHeadRightThinCell(row, "");
		this.otherElbSheet.addMergedRegion(new CellRangeAddress(row.getRowNum(), row.getRowNum(), 8, 15));
	}
  
	private void makeOtherELBTargetHealth(TargetHealthDescription targetHealthDescription) {
	  
		TargetDescription targetDescription = targetHealthDescription.target();
		TargetHealth targetHealth = targetHealthDescription.targetHealth();
		
		XSSFRow row = this.xssfHelper.createRow(this.otherElbSheet, 1);
		this.xssfHelper.setLeftThinCell(row, Integer.toString(row.getRowNum() - 1));
		this.xssfHelper.setCell(row, targetDescription.availabilityZone());
		this.xssfHelper.setCell(row, targetDescription.id());
		this.xssfHelper.setCell(row, targetDescription.port().toString());
		this.xssfHelper.setCell(row, targetHealthDescription.healthCheckPort());
		this.xssfHelper.setCell(row, targetHealth == null ? "" : this.getEnumName(targetHealth.state()));
		this.xssfHelper.setCell(row, targetHealth == null ? "" : this.getEnumName(targetHealth.reason()));
		this.xssfHelper.setCell(row, targetHealth == null ? "" : targetHealth.description());
		this.xssfHelper.setCell(row, "");
		this.xssfHelper.setCell(row, "");
		this.xssfHelper.setCell(row, "");
		this.xssfHelper.setCell(row, "");
		this.xssfHelper.setCell(row, "");
		this.xssfHelper.setCell(row, "");
		this.xssfHelper.setCell(row, "");
		this.xssfHelper.setRightThinCell(row, "");
		this.otherElbSheet.addMergedRegion(new CellRangeAddress(row.getRowNum(), row.getRowNum(), 8, 15));
	}
  
	private void makeOtherELBTargetGroupHeader() {
		XSSFRow row = this.xssfHelper.createRow(this.otherElbSheet, 1);
		this.xssfHelper.setLeftThinCell(row, Integer.toString(row.getRowNum() - 1));
		this.xssfHelper.setSecondSubHeadCell(row, "Action Order");
		this.xssfHelper.setSecondSubHeadCell(row, "Action Type");
		this.xssfHelper.setSecondSubHeadCell(row, "Target Group Name");
		this.xssfHelper.setSecondSubHeadCell(row, "Target Type");
		this.xssfHelper.setSecondSubHeadCell(row, "Target Protocol");
		this.xssfHelper.setSecondSubHeadCell(row, "Target Port");
		this.xssfHelper.setSecondSubHeadCell(row, "HealthCheck Protocol");
		this.xssfHelper.setSecondSubHeadCell(row, "HealthCheck Port");
		this.xssfHelper.setSecondSubHeadCell(row, "HealthCheck Path");
		this.xssfHelper.setSecondSubHeadCell(row, "HealthCheck Interval Seconds");
		this.xssfHelper.setSecondSubHeadCell(row, "HealthCheck Timeout Seconds");
		this.xssfHelper.setSecondSubHeadCell(row, "Healthy Threshold Count");
		this.xssfHelper.setSecondSubHeadCell(row, "Unhealthy Threshold Count");
		this.xssfHelper.setSecondSubHeadCell(row, "Target Group Arn");
		this.xssfHelper.setSecondSubHeadRightThinCell(row, "Target Group Attributes");
	}
  
	private void makeOtherELBTargetGroup(Action action, TargetGroup targetGroup, String targetGroupAttributes) {
		XSSFRow row = null;

		row = this.xssfHelper.createRow(this.otherElbSheet, 1);
		this.xssfHelper.setLeftThinCell(row, Integer.toString(row.getRowNum() - 1));
		this.xssfHelper.setCell(row, action.order() == null ? "" : action.order().toString());
		this.xssfHelper.setCell(row, this.getEnumName(action.type()));
		this.xssfHelper.setCell(row, targetGroup.targetGroupName());
		this.xssfHelper.setCell(row, this.getEnumName(targetGroup.targetType()));
		this.xssfHelper.setCell(row, this.getEnumName(targetGroup.protocol()));
		this.xssfHelper.setCell(row, targetGroup.port().toString());
		this.xssfHelper.setCell(row, this.getEnumName(targetGroup.healthCheckProtocol()));
		this.xssfHelper.setCell(row, targetGroup.healthCheckPort());
		this.xssfHelper.setCell(row, targetGroup.healthCheckPath());
		this.xssfHelper.setCell(row, targetGroup.healthCheckIntervalSeconds().toString());
		this.xssfHelper.setCell(row, targetGroup.healthCheckTimeoutSeconds().toString());
		this.xssfHelper.setCell(row, targetGroup.healthyThresholdCount().toString());
		this.xssfHelper.setCell(row, targetGroup.unhealthyThresholdCount().toString());
		this.xssfHelper.setCell(row, targetGroup.targetGroupArn());
		this.xssfHelper.setCell(row, targetGroupAttributes);
	}

	
	private void makeEC2Instance() {
		XSSFRow row = null;
		List<Reservation> reservations = this.amazonClients.ec2Client.describeInstances().reservations();
		for (Reservation reservation : reservations) {
			List<Instance> instances = reservation.instances();
			for (Instance instance : instances) {
				Subnet subnet = null;
				if(instance.subnetId() != null) {
					subnet = this.amazonClients.ec2Client
						.describeSubnets(
								DescribeSubnetsRequest.builder().subnetIds(instance.subnetId()).build())
						.subnets().get(0);
				}

				row = this.xssfHelper.createRow(this.ec2InstanceSheet, 1);
				this.xssfHelper.setLeftThinCell(row, Integer.toString(row.getRowNum() - 1));
				this.xssfHelper.setCell(row, getNameTagValue(instance.tags()));
				this.xssfHelper.setCell(row, instance.instanceId());
				this.xssfHelper.setCell(row, this.getEnumName(instance.instanceType()));
				this.xssfHelper.setCell(row, subnet == null ? "" : subnet.availabilityZone());
				this.xssfHelper.setCell(row, instance.privateIpAddress());
				this.xssfHelper.setCell(row, instance.privateDnsName());
				this.xssfHelper.setCell(row, instance.publicIpAddress());
				this.xssfHelper.setCell(row, instance.publicDnsName());

				StringBuffer securityGroups = new StringBuffer();
				List<GroupIdentifier> groupIdentifiers = instance.securityGroups();
				for (GroupIdentifier groupIdentifier : groupIdentifiers) {
					if (securityGroups.length() > 0) {
						securityGroups.append("\n");
					}
					securityGroups.append(groupIdentifier.groupName());
					securityGroups.append(" (");
					securityGroups.append(groupIdentifier.groupId());
					securityGroups.append(")");
				}
				this.xssfHelper.setCell(row, securityGroups.toString());

				this.xssfHelper.setCell(row, formatDate.format(instance.launchTime()));

				InstanceState instanceState = instance.state();
				this.xssfHelper.setCell(row,
						instanceState == null ? "" : instanceState.code() + "(" + instanceState.name() + ")");

				StateReason stateReason = instance.stateReason();
				this.xssfHelper.setCell(row,
						stateReason == null ? "" : stateReason.code() + "(" + stateReason.message() + ")");
				this.xssfHelper.setCell(row, instance.stateTransitionReason());

				Monitoring monitoring = instance.monitoring();
				this.xssfHelper.setCell(row, monitoring == null ? "" : this.getEnumName(monitoring.state()));

				this.xssfHelper.setCell(row, Boolean
						.toString(instance.sourceDestCheck() == null ? false : instance.sourceDestCheck()));
				this.xssfHelper.setCell(row, instance.spotInstanceRequestId());
				this.xssfHelper.setCell(row, instance.sriovNetSupport());

				this.xssfHelper.setCell(row, this.getEnumName(instance.virtualizationType()));
				this.xssfHelper.setCell(row, this.getEnumName(instance.platform()));
				this.xssfHelper.setCell(row, this.getEnumName(instance.architecture()));
				this.xssfHelper.setCell(row, instance.kernelId());
				this.xssfHelper.setCell(row,
						instance.enaSupport() == null ? "" : Boolean.toString(instance.enaSupport()));
				this.xssfHelper.setCell(row, this.getEnumName(instance.hypervisor()));
				this.xssfHelper.setCell(row, instance.clientToken());
				this.xssfHelper.setCell(row, Integer.toString(instance.amiLaunchIndex()));
				this.xssfHelper.setCell(row, instance.imageId());
				this.xssfHelper.setCell(row, this.getEnumName(instance.instanceLifecycle()));
				this.xssfHelper.setCell(row, instance.keyName());

				CpuOptions cpuOptions = instance.cpuOptions();
				this.xssfHelper.setCell(row,
						cpuOptions == null ? ""
								: "Core : " + cpuOptions.coreCount() + "\nThreads Per Core : "
										+ cpuOptions.threadsPerCore());

				StringBuffer gpuAssociation = new StringBuffer();
				List<ElasticGpuAssociation> elasticGpuAssociations = instance.elasticGpuAssociations();
				for (ElasticGpuAssociation elasticGpuAssociation : elasticGpuAssociations) {
					if (gpuAssociation.length() > 0) {
						gpuAssociation.append("\n");
					}
					gpuAssociation.append("AssociationID : " + elasticGpuAssociation.elasticGpuAssociationId());
					gpuAssociation.append(", GpuID : " + elasticGpuAssociation.elasticGpuId());
					gpuAssociation
							.append(", AssociationTime : " + elasticGpuAssociation.elasticGpuAssociationTime());
					gpuAssociation
							.append(", AssociationState : " + elasticGpuAssociation.elasticGpuAssociationState());
				}
				this.xssfHelper.setCell(row, gpuAssociation.toString());

				this.xssfHelper.setCell(row, Boolean.toString(instance.ebsOptimized()));
				this.xssfHelper.setCell(row, instance.ramdiskId());
				this.xssfHelper.setCell(row, instance.rootDeviceName());
				this.xssfHelper.setCell(row, this.getEnumName(instance.rootDeviceType()));

				StringBuffer blockDeviceMapping = new StringBuffer();
				List<InstanceBlockDeviceMapping> instanceBlockDeviceMappings = instance.blockDeviceMappings();
				for (InstanceBlockDeviceMapping instanceBlockDeviceMapping : instanceBlockDeviceMappings) {
					if (blockDeviceMapping.length() > 0) {
						blockDeviceMapping.append("\n");
					}
					blockDeviceMapping.append("DeviceName : " + instanceBlockDeviceMapping.deviceName());
					EbsInstanceBlockDevice ebsInstanceBlockDevice = instanceBlockDeviceMapping.ebs();
					if (ebsInstanceBlockDevice != null) {
						blockDeviceMapping.append(", VolumeId : " + ebsInstanceBlockDevice.volumeId());
						blockDeviceMapping
								.append(", Attach Time : " + formatDate.format(ebsInstanceBlockDevice.attachTime()));
						blockDeviceMapping.append(", DeleteOnTermination : "
								+ Boolean.toString(ebsInstanceBlockDevice.deleteOnTermination()));
						blockDeviceMapping.append(", Status : " + ebsInstanceBlockDevice.status());
					}
				}
				this.xssfHelper.setCell(row, blockDeviceMapping.toString());

				StringBuffer networkInterfaces = new StringBuffer();
				List<InstanceNetworkInterface> instanceNetworkInterfaces = instance.networkInterfaces();
				for (InstanceNetworkInterface instanceNetworkInterface : instanceNetworkInterfaces) {
					if (networkInterfaces.length() > 0) {
						networkInterfaces.append("\n");
					}
					networkInterfaces
							.append("NetworkInterfaceid : " + instanceNetworkInterface.networkInterfaceId());
					networkInterfaces.append(", OwnerId : " + instanceNetworkInterface.ownerId());
					networkInterfaces.append(", VpcId : " + instanceNetworkInterface.vpcId());
					networkInterfaces.append(", SubnetId : " + instanceNetworkInterface.subnetId());
					networkInterfaces.append(", Status : " + instanceNetworkInterface.status());
					networkInterfaces.append(", SourceDestCheck : " + instanceNetworkInterface.sourceDestCheck());
					networkInterfaces.append(", Mac : " + instanceNetworkInterface.macAddress());
					networkInterfaces
							.append(", Private Ip Address : " + instanceNetworkInterface.privateIpAddress());
					networkInterfaces.append(", Private Dns Name : " + instanceNetworkInterface.privateDnsName());

					List<InstancePrivateIpAddress> instancePrivateIpAddresses = instanceNetworkInterface
							.privateIpAddresses();
					networkInterfaces.append(", Private Ips [");
					for (InstancePrivateIpAddress instancePrivateIpAddress : instancePrivateIpAddresses) {
						networkInterfaces.append("[ Primary : " + instancePrivateIpAddress.primary());
						networkInterfaces.append(", Ip : " + instancePrivateIpAddress.privateIpAddress());
						networkInterfaces.append(", Dns Name : " + instancePrivateIpAddress.privateDnsName());
						networkInterfaces.append("], ");
					}
					networkInterfaces.append("]");

					List<InstanceIpv6Address> instanceIpv6Addresses = instanceNetworkInterface.ipv6Addresses();
					networkInterfaces.append(", Ipv6s [");
					for (InstanceIpv6Address instanceIpv6Address : instanceIpv6Addresses) {
						networkInterfaces.append(instanceIpv6Address.ipv6Address() + ", ");
					}
					networkInterfaces.append("]");

					InstanceNetworkInterfaceAssociation instanceNetworkInterfaceAssociation = instanceNetworkInterface
							.association();
					if (instanceNetworkInterfaceAssociation != null) {
						networkInterfaces.append("Association [");
						networkInterfaces.append("IpOwnerId : " + instanceNetworkInterfaceAssociation.ipOwnerId());
						networkInterfaces.append(", PublicIP : " + instanceNetworkInterfaceAssociation.publicIp());
						networkInterfaces
								.append(", PublicDnsName : " + instanceNetworkInterfaceAssociation.publicDnsName());
						networkInterfaces.append("]");
					}

					InstanceNetworkInterfaceAttachment instanceNetworkInterfaceAttachment = instanceNetworkInterface
							.attachment();
					if (instanceNetworkInterfaceAttachment != null) {
						networkInterfaces.append(", Attachment [");
						networkInterfaces
								.append("AttachmentId : " + instanceNetworkInterfaceAttachment.attachmentId());
						networkInterfaces.append(", AttachTime : "
								+ formatDate.format(instanceNetworkInterfaceAttachment.attachTime()));
						networkInterfaces.append(", DeleteOnTermination : "
								+ Boolean.toString(instanceNetworkInterfaceAttachment.deleteOnTermination()));
						networkInterfaces.append(", DeviceIndex : "
								+ Integer.toString(instanceNetworkInterfaceAttachment.deviceIndex()));
						networkInterfaces.append(", Status : " + instanceNetworkInterfaceAttachment.status());
						networkInterfaces.append("]");
					}

					networkInterfaces.append(", Description : " + instanceNetworkInterface.description());
					List<GroupIdentifier> ngroupIdentifiers = instanceNetworkInterface.groups();
					networkInterfaces.append(", Groups [");
					for (GroupIdentifier ngroupIdentifier : ngroupIdentifiers) {
						networkInterfaces
								.append(ngroupIdentifier.groupId() + "(" + ngroupIdentifier.groupName() + ")");
						networkInterfaces.append(", ");
					}
					networkInterfaces.append("]");

				}
				this.xssfHelper.setCell(row, networkInterfaces.toString());

				Placement placement = instance.placement();
				this.xssfHelper.setCell(row,
						placement == null ? ""
								: "Affinify : " + placement.affinity() + "\nAvailabilityZone : "
										+ placement.availabilityZone() + "\nGroupName : " + placement.groupName()
										+ "\nHostId : " + placement.hostId() + "\nSpreadDomain : "
										+ placement.spreadDomain() + "\nTenancy : " + placement.tenancy());

				StringBuffer productCds = new StringBuffer();
				List<ProductCode> productCodes = instance.productCodes();
				for (ProductCode productCode : productCodes) {
					if (productCds.length() > 0) {
						productCds.append("\n");
					}
					productCds.append("ID : " + productCode.productCodeId());
					productCds.append(", TYPE : " + productCode.productCodeType());
				}
				this.xssfHelper.setCell(row, productCds.toString());

				IamInstanceProfile iamInstanceProfile = instance.iamInstanceProfile();
				this.xssfHelper.setCell(row, iamInstanceProfile == null ? ""
						: "ID : " + iamInstanceProfile.id() + ", ARN : " + iamInstanceProfile.arn());

				this.xssfHelper.setRightThinCell(row, this.getAllTagValue(instance.tags()));
			}
		}

	}

	private void makeSecurityGroup() {
		XSSFRow row = null;

		int securityGroupStartRowNum = this.securityGroupSheet.getLastRowNum();

		List<SecurityGroup> securityGroups = this.amazonClients.ec2Client.describeSecurityGroups().securityGroups();
		for (SecurityGroup securityGroup : securityGroups) {
			row = this.xssfHelper.createRow(this.securityGroupSheet, 1);
			this.xssfHelper.setLeftThinCell(row, Integer.toString(row.getRowNum() - 1));
			this.xssfHelper.setCell(row, this.region.id());
			this.xssfHelper.setCell(row, securityGroup.vpcId());
			this.xssfHelper.setBrownBoldLeftThinCell(row, securityGroup.groupId() + "["
					+ securityGroup.groupName() + "] - " + securityGroup.description());
			this.xssfHelper.setCell(row, "");
			this.xssfHelper.setCell(row, "");
			this.xssfHelper.setCell(row, "");
			this.xssfHelper.setCell(row, "");
			this.xssfHelper.setBrownBoldLeftThinCell(row, this.getAllTagValue(securityGroup.tags()));
			this.xssfHelper.setCell(row, "");
			this.xssfHelper.setCell(row, "");
			this.xssfHelper.setCell(row, "");
			this.xssfHelper.setRightThinCell(row, "");

			this.securityGroupSheet.addMergedRegion(new CellRangeAddress(row.getRowNum(), row.getRowNum(), 3, 7));
			this.securityGroupSheet.addMergedRegion(new CellRangeAddress(row.getRowNum(), row.getRowNum(), 8, 12));

			row = this.xssfHelper.createRow(this.securityGroupSheet, 1);
			this.xssfHelper.setLeftThinCell(row, Integer.toString(row.getRowNum() - 1));
			this.xssfHelper.setCell(row, "");
			this.xssfHelper.setCell(row, "");
			this.xssfHelper.setSubHeadLeftThinCell(row, "Inbound");
			this.xssfHelper.setCell(row, "");
			this.xssfHelper.setCell(row, "");
			this.xssfHelper.setCell(row, "");
			this.xssfHelper.setCell(row, "");
			this.xssfHelper.setSubHeadLeftThinCell(row, "Outbound");
			this.xssfHelper.setCell(row, "");
			this.xssfHelper.setCell(row, "");
			this.xssfHelper.setCell(row, "");
			this.xssfHelper.setRightThinCell(row, "");

			this.securityGroupSheet.addMergedRegion(new CellRangeAddress(row.getRowNum(), row.getRowNum(), 3, 7));
			this.securityGroupSheet.addMergedRegion(new CellRangeAddress(row.getRowNum(), row.getRowNum(), 8, 12));

			row = this.xssfHelper.createRow(this.securityGroupSheet, 1);
			this.xssfHelper.setLeftThinCell(row, Integer.toString(row.getRowNum() - 1));
			this.xssfHelper.setCell(row, "");
			this.xssfHelper.setCell(row, "");
			this.xssfHelper.setSecondSubHeadLeftThinCell(row, "Type");
			this.xssfHelper.setSecondSubHeadCell(row, "Protocol");
			this.xssfHelper.setSecondSubHeadCell(row, "Port Range");
			this.xssfHelper.setSecondSubHeadCell(row, "Source");
			this.xssfHelper.setSecondSubHeadCell(row, "Description");
			this.xssfHelper.setSecondSubHeadLeftThinCell(row, "Type");
			this.xssfHelper.setSecondSubHeadCell(row, "Protocol");
			this.xssfHelper.setSecondSubHeadCell(row, "PortRange");
			this.xssfHelper.setSecondSubHeadCell(row, "Source");
			this.xssfHelper.setSecondSubHeadRightThinCell(row, "Description");

			List<IpPermission> inbounds = securityGroup.ipPermissions();
			List<IpPermission> outbounds = securityGroup.ipPermissionsEgress();

			int inboundSize = getIpPermissionSize(inbounds);
			int outboundSize = getIpPermissionSize(outbounds);

			int startRowNum = row.getRowNum() + 1;
			int rowNum = inboundSize > outboundSize ? inboundSize : outboundSize;
			for (int iRowNum = 0; iRowNum < rowNum; iRowNum++) {
				row = this.xssfHelper.createRow(this.securityGroupSheet, 1);
				this.xssfHelper.setLeftThinCell(row, Integer.toString(row.getRowNum() - 1));
				this.xssfHelper.setCell(row, "");
				this.xssfHelper.setCell(row, "");
				this.xssfHelper.setLeftThinCell(row, "");
				this.xssfHelper.setCell(row, "");
				this.xssfHelper.setCell(row, "");
				this.xssfHelper.setCell(row, "");
				this.xssfHelper.setCell(row, "");
				this.xssfHelper.setLeftThinCell(row, "");
				this.xssfHelper.setCell(row, "");
				this.xssfHelper.setCell(row, "");
				this.xssfHelper.setCell(row, "");
				this.xssfHelper.setRightThinCell(row, "");
			}
			this.xssfHelper.setSubLastRowStyle(row);
			if (rowNum > 0) {
				row = getIpPermissionList(inbounds, startRowNum, false);
				row = getIpPermissionList(outbounds, startRowNum, true);
			}
		}
		this.securityGroupSheet.addMergedRegion(
				new CellRangeAddress(securityGroupStartRowNum + 1, this.securityGroupSheet.getLastRowNum(), 2, 2));
		if (this.securityGroupSheet.getLastRowNum() > 2) {
			this.securityGroupSheet
					.addMergedRegion(new CellRangeAddress(2, this.securityGroupSheet.getLastRowNum(), 1, 1));
		}
	}

	private int getIpPermissionSize(List<IpPermission> ipPermissions) {
		int IpPermissionsSize = 0;
		for (IpPermission ipPermission : ipPermissions) {
			IpPermissionsSize += ipPermission.prefixListIds().size();
			IpPermissionsSize += ipPermission.ipRanges().size();
			IpPermissionsSize += ipPermission.ipv6Ranges().size();
			IpPermissionsSize += ipPermission.userIdGroupPairs().size();
		}
		return IpPermissionsSize;
	}

	private String getSecurityGroupRuleType(String ipProtocol, Integer fromPort, Integer toPort) {
		String type = ("icmp".equals(ipProtocol)) && (fromPort != null && fromPort.intValue() == -1) ? "All "
				: (fromPort != null && fromPort.intValue() == 0) && (toPort != null && toPort.intValue() == 65535)
						? "All "
						: ("-1".equals(ipProtocol)) && (fromPort == null) ? "All traffic" : "Custom ";

		String stype = "";
		if ((fromPort != null) && (fromPort.compareTo(toPort) == 0)) {
			switch (fromPort.intValue()) {
			case 22:
				stype = "SSH";
				break;
			case 80:
				stype = "HTTP";
				break;
			case 443:
				stype = "HTTPS";
				break;
			case 3306:
				stype = "MYSQL/Aurora";
				break;
			case 1521:
				stype = "Oracle-RDS";
				break;
			case 25:
				stype = "SMTP";
				break;
			case 53:
				if ("udp".equals(ipProtocol)) {
					stype = "DNS (UDP)";
				}
				if ("tcp".equals(ipProtocol)) {
					stype = "DNS (TCP)";
				}
				break;
			case 110:
				stype = "POP3";
				break;
			case 143:
				stype = "IMAP";
				break;
			case 389:
				stype = "LDAP";
				break;
			case 465:
				stype = "SMTPS";
				break;
			case 993:
				stype = "IMAPS";
				break;
			case 995:
				stype = "POP3S";
				break;
			case 1443:
				stype = "MS SQL";
				break;
			case 2049:
				stype = "NFS";
				break;
			case 3389:
				stype = "RDP";
				break;
			case 5439:
				stype = "Redshift";
				break;
			case 5432:
				stype = "PostgreSQL";
				break;
			}
		}
		StringBuffer typeBuffer = new StringBuffer();
		if ("".equals(stype)) {
			typeBuffer.append(type);
			typeBuffer.append("-1".equals(ipProtocol) ? "" : ipProtocol.toUpperCase());
			if ("Custom ".equals(type)) {
				typeBuffer.append(" Rule");
			}
		} else {
			typeBuffer.append(stype);
		}
		return typeBuffer.toString();
	}

	private String getPortRange(Integer fromPort, Integer toPort) {
		StringBuffer portRange = new StringBuffer();
		portRange.append((fromPort != null && fromPort.intValue() == -1) ? "N/A" : fromPort == null ? "All" : fromPort);
		if ((fromPort != null) && (fromPort.compareTo(toPort) != 0)) {
			portRange.append(" - ");
			portRange.append(toPort);
		}
		return portRange.toString();
	}

	private XSSFRow getIpPermissionCom(IpPermission ipPermission, int rowNum, String source, String description,
			boolean isEngress) {
		XSSFRow row = this.securityGroupSheet.getRow(rowNum);

		this.xssfHelper.setCell(row, 3 + (isEngress ? 5 : 0), getSecurityGroupRuleType(ipPermission.ipProtocol(),
				ipPermission.fromPort(), ipPermission.toPort()));
		this.xssfHelper.setCell(row, 4 + (isEngress ? 5 : 0),
				("-1".equals(ipPermission.ipProtocol())) || (("icmp".equals(ipPermission.ipProtocol()))
						&& (ipPermission.fromPort() != null) && (ipPermission.fromPort().intValue() == -1))
								? "All"
								: ipPermission.ipProtocol().toUpperCase());
		this.xssfHelper.setCell(row, 5 + (isEngress ? 5 : 0),
				getPortRange(ipPermission.fromPort(), ipPermission.toPort()));
		this.xssfHelper.setCell(row, 6 + (isEngress ? 5 : 0), source);
		this.xssfHelper.setCell(row, 7 + (isEngress ? 5 : 0), description);

		return row;
	}

	private String getUserIdGroupPair(String groupId) {
		StringBuffer userIdGroupPair = new StringBuffer();
		userIdGroupPair.append(groupId);
		userIdGroupPair.append(" (");
		userIdGroupPair.append(getSGName(groupId));
		userIdGroupPair.append(")");
		return userIdGroupPair.toString();
	}

	private XSSFRow getIpPermissionList(List<IpPermission> ipPermissions, int startRowNum, boolean isEngress) {
		XSSFRow row = null;
		int rowNum = 0;
		for (IpPermission ipPermission : ipPermissions) {

			List<PrefixListId> prefixListIds = ipPermission.prefixListIds();
			for (PrefixListId prefixListId : prefixListIds) {
				row = getIpPermissionCom(ipPermission, startRowNum + rowNum, prefixListId.prefixListId(),
						prefixListId.description(), isEngress);
				rowNum++;
			}
			List<IpRange> ipRanges = ipPermission.ipRanges();
			for (IpRange ipRange : ipRanges) {
				row = getIpPermissionCom(ipPermission, startRowNum + rowNum, ipRange.cidrIp(),
						ipRange.description(), isEngress);
				rowNum++;
			}
			List<Ipv6Range> ipv6Ranges = ipPermission.ipv6Ranges();
			for (Ipv6Range ipv6Range : ipv6Ranges) {
				row = getIpPermissionCom(ipPermission, startRowNum + rowNum, ipv6Range.cidrIpv6(),
						ipv6Range.description(), isEngress);
				rowNum++;
			}
			List<UserIdGroupPair> userIdGroupPairs = ipPermission.userIdGroupPairs();
			for (UserIdGroupPair userIdGroupPair : userIdGroupPairs) {
				row = getIpPermissionCom(ipPermission, startRowNum + rowNum,
						getUserIdGroupPair(userIdGroupPair.groupId()), userIdGroupPair.description(), isEngress);
				rowNum++;
			}
		}
		return row;
	}

	private String getSGName(String GroupId) {
		String sgName = "";
		try {
			DescribeSecurityGroupsResponse describeSecurityGroupsResult = this.amazonClients.ec2Client
					.describeSecurityGroups(DescribeSecurityGroupsRequest.builder().groupIds(new String[] { GroupId }).build());
			sgName = ((SecurityGroup) describeSecurityGroupsResult.securityGroups().get(0)).groupName();
		} catch (Exception localException) {
		}
		return sgName;
	}

	
	private void makeInternetGateway() {
		XSSFRow row = null;
		List<InternetGateway> internetGateways = this.amazonClients.ec2Client
				.describeInternetGateways().internetGateways();
		for (int iIGW = 0; iIGW < internetGateways.size(); iIGW++) {
			InternetGateway internetGateway = internetGateways.get(iIGW);

			row = this.xssfHelper.createRow(this.internetGatewaySheet, 1);
			this.xssfHelper.setLeftThinCell(row, Integer.toString(row.getRowNum() - 1));
			this.xssfHelper.setCell(row, this.region.id());

			StringBuffer vpcBuffer = new StringBuffer();
			List<InternetGatewayAttachment> internetGatewayAttachments = internetGateway.attachments();
			for (InternetGatewayAttachment internetGatewayAttachment : internetGatewayAttachments) {
				if (vpcBuffer.length() > 0) {
					vpcBuffer.append("\n");
				}
				vpcBuffer.append(internetGatewayAttachment.vpcId());
				vpcBuffer.append("(");
				vpcBuffer.append(internetGatewayAttachment.state());
				vpcBuffer.append(")");
			}
			this.xssfHelper.setCell(row, vpcBuffer.toString());
			this.xssfHelper.setCell(row, getNameTagValue(internetGateway.tags()));
			this.xssfHelper.setCell(row, internetGateway.internetGatewayId());
			this.xssfHelper.setRightThinCell(row, this.getAllTagValue(internetGateway.tags()));
		}
	}

	
	private HashMap<String, String> makeRouteTable() {
		XSSFRow row = null;

		HashMap<String, String> vpcMainRouteTables = new HashMap<String, String>();
		List<RouteTable> routeTables = this.amazonClients.ec2Client.describeRouteTables().routeTables();
		for (int i = 0; i < routeTables.size(); i++) {
			boolean isMain = false;
			RouteTable routeTable = routeTables.get(i);

			List<RouteTableAssociation> routeTableAssociations = routeTable.associations();
			for (RouteTableAssociation routeTableAssociation : routeTableAssociations) {
				if (routeTableAssociation.main().booleanValue()) {
					isMain = true;
					vpcMainRouteTables.put(routeTable.vpcId(),
							routeTable.routeTableId() + " | " + getNameTagValue(routeTable.tags()));
				}
			}
			row = this.xssfHelper.createRow(this.routeTableSheet, 1);
			int firstRowNum = row.getRowNum();
			this.xssfHelper.setLeftThinCell(row, Integer.toString(row.getRowNum() - 1));
			this.xssfHelper.setCell(row, this.region.id());
			this.xssfHelper.setCell(row, routeTable.vpcId());
			this.xssfHelper.setBrownBoldLeftThinCell(row, getNameTagValue(routeTable.tags()));
			this.xssfHelper.setCell(row, routeTable.routeTableId());
			this.xssfHelper.setLeftRightThinCell(row, isMain ? "YES" : "NO");
			this.xssfHelper.setRightThinCell(row, "");
			this.xssfHelper.setCell(row, this.getAllTagValue(routeTable.tags()));
			this.xssfHelper.setRightThinCell(row, "");

			this.routeTableSheet.addMergedRegion(new CellRangeAddress(row.getRowNum(), row.getRowNum(), 5, 6));
			this.routeTableSheet.addMergedRegion(new CellRangeAddress(row.getRowNum(), row.getRowNum(), 7, 8));

			row = this.xssfHelper.createRow(this.routeTableSheet, 1);
			this.xssfHelper.setLeftThinCell(row, "");
			this.xssfHelper.setCell(row, "");
			this.xssfHelper.setCell(row, "");
			this.xssfHelper.setSubHeadLeftThinCell(row, "Routes");
			this.xssfHelper.setSubHeadCell(row, "");
			this.xssfHelper.setSubHeadCell(row, "");
			this.xssfHelper.setSubHeadCell(row, "");
			this.xssfHelper.setSubHeadLeftThinCell(row, "Associated Subnet");
			this.xssfHelper.setSubHeadRightThinCell(row, "");

			row = this.xssfHelper.createRow(this.routeTableSheet, 1);
			this.xssfHelper.setLeftThinCell(row, "");
			this.xssfHelper.setCell(row, "");
			this.xssfHelper.setCell(row, "");
			this.xssfHelper.setSecondSubHeadLeftThinCell(row, "Destination");
			this.xssfHelper.setSecondSubHeadCell(row, "Target");
			this.xssfHelper.setSecondSubHeadCell(row, "State");
			this.xssfHelper.setSecondSubHeadCell(row, "Origin");
			this.xssfHelper.setSecondSubHeadLeftThinCell(row, "Subnet");
			this.xssfHelper.setSecondSubHeadRightThinCell(row, "CIDR");

			List<Integer> rows = new ArrayList<Integer>();

			List<Route> routes = routeTable.routes();
			for (int iRoute = 0; iRoute < routes.size(); iRoute++) {
				Route route = routes.get(iRoute);

				String destination = "";
				String target = "";
				if (route.destinationCidrBlock() != null) {
					destination = route.destinationCidrBlock();
				} else if (route.destinationIpv6CidrBlock() != null) {
					destination = route.destinationIpv6CidrBlock();
				} else if (route.destinationPrefixListId() != null) {
					destination = route.destinationPrefixListId();
				}
				if (route.egressOnlyInternetGatewayId() != null) {
					target = route.egressOnlyInternetGatewayId();
				} else if (route.gatewayId() != null) {
					target = route.gatewayId();
				} else if (route.natGatewayId() != null) {
					target = route.natGatewayId();
				} else if (route.networkInterfaceId() != null) {
					target = route.networkInterfaceId();
				} else if (route.vpcPeeringConnectionId() != null) {
					target = route.vpcPeeringConnectionId();
				}
				row = this.xssfHelper.createRow(this.routeTableSheet, 1);
				this.xssfHelper.setLeftThinCell(row, "");
				this.xssfHelper.setCell(row, "");
				this.xssfHelper.setCell(row, "");
				this.xssfHelper.setLeftThinCell(row, destination);
				this.xssfHelper.setCell(row, target);
				this.xssfHelper.setCell(row, this.getEnumName(route.state()));
				this.xssfHelper.setCell(row, this.getEnumName(route.origin()));
				this.xssfHelper.setLeftThinCell(row, "");
				this.xssfHelper.setRightThinCell(row, "");

				rows.add(Integer.valueOf(row.getRowNum()));
			}
			int lastRowNum = row.getRowNum();

			int iSubnetCnt = 0;
			for (int iAssoc = 0; iAssoc < routeTableAssociations.size(); iAssoc++) {
				RouteTableAssociation routeTableAssociation = routeTableAssociations.get(iAssoc);

				if (!routeTableAssociation.main()) {
					List<Subnet> subnets = this.amazonClients.ec2Client.describeSubnets(DescribeSubnetsRequest.builder()
							.subnetIds(new String[] { routeTableAssociation.subnetId() }).build()).subnets();
					for (int iSubnet = 0; iSubnet < subnets.size(); iSubnet++) {
						Subnet subnet = subnets.get(iSubnet);
						if (rows.size() > iSubnetCnt) {
							row = this.routeTableSheet.getRow(rows.get(iSubnetCnt).intValue());
						} else {
							row = null;
						}
						if (row == null) {
							row = this.xssfHelper.createRow(this.routeTableSheet, 1);
							this.xssfHelper.setLeftThinCell(row, "");
							this.xssfHelper.setCell(row, "");
							this.xssfHelper.setCell(row, "");
							this.xssfHelper.setLeftThinCell(row, "");
							this.xssfHelper.setCell(row, "");
							this.xssfHelper.setCell(row, "");
							this.xssfHelper.setCell(row, "");
							this.xssfHelper.setLeftThinCell(row,
									subnet.subnetId() + " | " + getNameTagValue(subnet.tags()));
							this.xssfHelper.setRightThinCell(row, subnet.cidrBlock());
						} else {
							row.getCell(row.getLastCellNum() - 2)
									.setCellValue(subnet.subnetId() + " | " + getNameTagValue(subnet.tags()));
							row.getCell(row.getLastCellNum() - 1).setCellValue(subnet.cidrBlock());
						}
						iSubnetCnt++;
					}
				}
			}

			if (lastRowNum < row.getRowNum()) {
				lastRowNum = row.getRowNum();
			}
			this.xssfHelper.setSubLastRowStyle(this.routeTableSheet.getRow(lastRowNum));
			if (firstRowNum > lastRowNum) {
				this.routeTableSheet.addMergedRegion(new CellRangeAddress(firstRowNum, lastRowNum, 0, 0));
				this.routeTableSheet.addMergedRegion(new CellRangeAddress(firstRowNum, lastRowNum, 1, 1));
				this.routeTableSheet.addMergedRegion(new CellRangeAddress(firstRowNum, lastRowNum, 2, 2));
			}
			this.routeTableSheet.addMergedRegion(new CellRangeAddress(firstRowNum + 1, firstRowNum + 1, 3, 6));
			this.routeTableSheet.addMergedRegion(new CellRangeAddress(firstRowNum + 1, firstRowNum + 1, 7, 8));
		}
		return vpcMainRouteTables;
	}

	private void makeSubnet() {
		XSSFRow row = null;

		List<Subnet> subnets = this.amazonClients.ec2Client.describeSubnets().subnets();
		for (int iSubnet = 0; iSubnet < subnets.size(); iSubnet++) {
			Subnet subnet = subnets.get(iSubnet);

			List<String> subnetList = new ArrayList<String>();
			subnetList.add(subnet.subnetId());

			List<RouteTable> routeTables = this.amazonClients.ec2Client
					.describeRouteTables(DescribeRouteTablesRequest.builder()
							.filters(Filter.builder().name("association.subnet-id").values(subnetList).build())
							.build())
					.routeTables();
			List<NetworkAcl> networkAcls = this.amazonClients.ec2Client
					.describeNetworkAcls(
							DescribeNetworkAclsRequest.builder()
								.filters(Filter.builder().name("association.subnet-id").values(subnetList).build())
								.build()
					)
					.networkAcls();

			String routeTableInfo = "";
			String networkAclInfo = "";
			if (routeTables.size() > 0) {
				RouteTable routeTable = routeTables.get(0);
				routeTableInfo = routeTable.routeTableId() + " | " + getNameTagValue(routeTable.tags());
			}
			if (networkAcls.size() > 0) {
				NetworkAcl networkAcl = networkAcls.get(0);
				networkAclInfo = networkAcl.networkAclId();
			}
			row = this.xssfHelper.createRow(this.subnetSheet, 1);
			this.xssfHelper.setLeftThinCell(row, Integer.toString(row.getRowNum() - 1));
			this.xssfHelper.setCell(row, this.region.id());
			this.xssfHelper.setCell(row, subnet.vpcId());
			this.xssfHelper.setCell(row, getNameTagValue(subnet.tags()));
			this.xssfHelper.setCell(row, subnet.subnetId());
			this.xssfHelper.setCell(row, subnet.cidrBlock());
			this.xssfHelper.setCell(row, subnet.availabilityZone());
			this.xssfHelper.setCell(row, routeTableInfo);
			this.xssfHelper.setCell(row, networkAclInfo);

			this.xssfHelper.setCell(row, Boolean.toString(subnet.assignIpv6AddressOnCreation()));
			this.xssfHelper.setCell(row, Integer.toString(subnet.availableIpAddressCount()));
			this.xssfHelper.setCell(row, Boolean.toString(subnet.defaultForAz()));

			StringBuffer subnetIpv6Cidr = new StringBuffer();
			List<SubnetIpv6CidrBlockAssociation> subnetIpv6CidrBlockAssociations = subnet
					.ipv6CidrBlockAssociationSet();
			for (SubnetIpv6CidrBlockAssociation subnetIpv6CidrBlockAssociation : subnetIpv6CidrBlockAssociations) {
				subnetIpv6Cidr.append("AssociationId : " + subnetIpv6CidrBlockAssociation.associationId());
				subnetIpv6Cidr.append("\n");
				subnetIpv6Cidr.append("CIDR Block : " + subnetIpv6CidrBlockAssociation.ipv6CidrBlock());
				subnetIpv6Cidr.append("\n");
				SubnetCidrBlockState subnetCidrBlockState = subnetIpv6CidrBlockAssociation.ipv6CidrBlockState();
				if (subnetCidrBlockState != null) {
					subnetIpv6Cidr.append("State : " + this.getEnumName(subnetCidrBlockState.state()));
					subnetIpv6Cidr.append("\n");
					subnetIpv6Cidr.append("Status Message : " + subnetCidrBlockState.statusMessage());
				}
			}
			this.xssfHelper.setCell(row, subnetIpv6Cidr.toString());
			this.xssfHelper.setCell(row, Boolean.toString(subnet.mapPublicIpOnLaunch()));
			this.xssfHelper.setCell(row, this.getEnumName(subnet.state()));
			this.xssfHelper.setCell(row, this.getAllTagValue(subnet.tags()));

			this.xssfHelper.setRightThinCell(row, "");

		}
	}
	
	private String getEnumName(Enum<?> enum1) {
		return enum1 == null ? "" : enum1.name();
	}

}
