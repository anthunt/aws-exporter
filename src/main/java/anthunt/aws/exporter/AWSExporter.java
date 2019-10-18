package anthunt.aws.exporter;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.text.DateFormat;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Date;
import java.util.HashMap;
import java.util.Iterator;
import java.util.List;
import java.util.Map;
import java.util.Set;

import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import com.amazonaws.regions.Regions;
import com.amazonaws.services.certificatemanager.model.CertificateDetail;
import com.amazonaws.services.certificatemanager.model.CertificateSummary;
import com.amazonaws.services.certificatemanager.model.DescribeCertificateRequest;
import com.amazonaws.services.certificatemanager.model.DescribeCertificateResult;
import com.amazonaws.services.certificatemanager.model.DomainValidation;
import com.amazonaws.services.certificatemanager.model.ExtendedKeyUsage;
import com.amazonaws.services.certificatemanager.model.KeyUsage;
import com.amazonaws.services.certificatemanager.model.ListCertificatesRequest;
import com.amazonaws.services.certificatemanager.model.ListCertificatesResult;
import com.amazonaws.services.certificatemanager.model.RenewalSummary;
import com.amazonaws.services.directconnect.model.BGPPeer;
import com.amazonaws.services.directconnect.model.Connection;
import com.amazonaws.services.directconnect.model.DescribeConnectionsResult;
import com.amazonaws.services.directconnect.model.DescribeDirectConnectGatewaysRequest;
import com.amazonaws.services.directconnect.model.DescribeDirectConnectGatewaysResult;
import com.amazonaws.services.directconnect.model.DescribeLagsRequest;
import com.amazonaws.services.directconnect.model.DescribeLagsResult;
import com.amazonaws.services.directconnect.model.DescribeLocationsResult;
import com.amazonaws.services.directconnect.model.DescribeVirtualGatewaysResult;
import com.amazonaws.services.directconnect.model.DescribeVirtualInterfacesResult;
import com.amazonaws.services.directconnect.model.DirectConnectGateway;
import com.amazonaws.services.directconnect.model.Lag;
import com.amazonaws.services.directconnect.model.Location;
import com.amazonaws.services.directconnect.model.RouteFilterPrefix;
import com.amazonaws.services.directconnect.model.VirtualGateway;
import com.amazonaws.services.directconnect.model.VirtualInterface;
import com.amazonaws.services.directory.model.ConditionalForwarder;
import com.amazonaws.services.directory.model.DescribeConditionalForwardersRequest;
import com.amazonaws.services.directory.model.DescribeConditionalForwardersResult;
import com.amazonaws.services.directory.model.DescribeDirectoriesResult;
import com.amazonaws.services.directory.model.DescribeDomainControllersRequest;
import com.amazonaws.services.directory.model.DescribeDomainControllersResult;
import com.amazonaws.services.directory.model.DescribeTrustsRequest;
import com.amazonaws.services.directory.model.DescribeTrustsResult;
import com.amazonaws.services.directory.model.DirectoryConnectSettingsDescription;
import com.amazonaws.services.directory.model.DirectoryDescription;
import com.amazonaws.services.directory.model.DirectoryVpcSettingsDescription;
import com.amazonaws.services.directory.model.DomainController;
import com.amazonaws.services.directory.model.InvalidParameterException;
import com.amazonaws.services.directory.model.RadiusSettings;
import com.amazonaws.services.directory.model.Trust;
import com.amazonaws.services.ec2.model.CidrBlock;
import com.amazonaws.services.ec2.model.CpuOptions;
import com.amazonaws.services.ec2.model.CustomerGateway;
import com.amazonaws.services.ec2.model.DescribeCustomerGatewaysResult;
import com.amazonaws.services.ec2.model.DescribeEgressOnlyInternetGatewaysRequest;
import com.amazonaws.services.ec2.model.DescribeEgressOnlyInternetGatewaysResult;
import com.amazonaws.services.ec2.model.DescribeInstancesRequest;
import com.amazonaws.services.ec2.model.DescribeInternetGatewaysRequest;
import com.amazonaws.services.ec2.model.DescribeNatGatewaysRequest;
import com.amazonaws.services.ec2.model.DescribeNatGatewaysResult;
import com.amazonaws.services.ec2.model.DescribeNetworkAclsRequest;
import com.amazonaws.services.ec2.model.DescribeRouteTablesRequest;
import com.amazonaws.services.ec2.model.DescribeSecurityGroupsRequest;
import com.amazonaws.services.ec2.model.DescribeSecurityGroupsResult;
import com.amazonaws.services.ec2.model.DescribeSubnetsRequest;
import com.amazonaws.services.ec2.model.DescribeVpcPeeringConnectionsResult;
import com.amazonaws.services.ec2.model.DescribeVpnConnectionsResult;
import com.amazonaws.services.ec2.model.DescribeVpnGatewaysResult;
import com.amazonaws.services.ec2.model.EbsInstanceBlockDevice;
import com.amazonaws.services.ec2.model.EgressOnlyInternetGateway;
import com.amazonaws.services.ec2.model.ElasticGpuAssociation;
import com.amazonaws.services.ec2.model.Filter;
import com.amazonaws.services.ec2.model.GroupIdentifier;
import com.amazonaws.services.ec2.model.IamInstanceProfile;
import com.amazonaws.services.ec2.model.InstanceBlockDeviceMapping;
import com.amazonaws.services.ec2.model.InstanceIpv6Address;
import com.amazonaws.services.ec2.model.InstanceNetworkInterface;
import com.amazonaws.services.ec2.model.InstanceNetworkInterfaceAssociation;
import com.amazonaws.services.ec2.model.InstanceNetworkInterfaceAttachment;
import com.amazonaws.services.ec2.model.InstancePrivateIpAddress;
import com.amazonaws.services.ec2.model.InstanceState;
import com.amazonaws.services.ec2.model.InternetGateway;
import com.amazonaws.services.ec2.model.InternetGatewayAttachment;
import com.amazonaws.services.ec2.model.IpPermission;
import com.amazonaws.services.ec2.model.IpRange;
import com.amazonaws.services.ec2.model.Ipv6CidrBlock;
import com.amazonaws.services.ec2.model.Ipv6Range;
import com.amazonaws.services.ec2.model.Monitoring;
import com.amazonaws.services.ec2.model.NatGateway;
import com.amazonaws.services.ec2.model.NatGatewayAddress;
import com.amazonaws.services.ec2.model.NetworkAcl;
import com.amazonaws.services.ec2.model.Placement;
import com.amazonaws.services.ec2.model.PrefixListId;
import com.amazonaws.services.ec2.model.ProductCode;
import com.amazonaws.services.ec2.model.ProvisionedBandwidth;
import com.amazonaws.services.ec2.model.Reservation;
import com.amazonaws.services.ec2.model.Route;
import com.amazonaws.services.ec2.model.RouteTable;
import com.amazonaws.services.ec2.model.RouteTableAssociation;
import com.amazonaws.services.ec2.model.SecurityGroup;
import com.amazonaws.services.ec2.model.StateReason;
import com.amazonaws.services.ec2.model.Subnet;
import com.amazonaws.services.ec2.model.SubnetCidrBlockState;
import com.amazonaws.services.ec2.model.SubnetIpv6CidrBlockAssociation;
import com.amazonaws.services.ec2.model.Tag;
import com.amazonaws.services.ec2.model.UserIdGroupPair;
import com.amazonaws.services.ec2.model.VgwTelemetry;
import com.amazonaws.services.ec2.model.Volume;
import com.amazonaws.services.ec2.model.VolumeAttachment;
import com.amazonaws.services.ec2.model.Vpc;
import com.amazonaws.services.ec2.model.VpcAttachment;
import com.amazonaws.services.ec2.model.VpcCidrBlockAssociation;
import com.amazonaws.services.ec2.model.VpcCidrBlockState;
import com.amazonaws.services.ec2.model.VpcIpv6CidrBlockAssociation;
import com.amazonaws.services.ec2.model.VpcPeeringConnection;
import com.amazonaws.services.ec2.model.VpcPeeringConnectionOptionsDescription;
import com.amazonaws.services.ec2.model.VpcPeeringConnectionStateReason;
import com.amazonaws.services.ec2.model.VpcPeeringConnectionVpcInfo;
import com.amazonaws.services.ec2.model.VpnConnection;
import com.amazonaws.services.ec2.model.VpnGateway;
import com.amazonaws.services.ec2.model.VpnStaticRoute;
import com.amazonaws.services.elasticache.model.CacheCluster;
import com.amazonaws.services.elasticache.model.DescribeCacheClustersRequest;
import com.amazonaws.services.elasticache.model.NodeGroup;
import com.amazonaws.services.elasticache.model.NodeGroupMember;
import com.amazonaws.services.elasticache.model.ReplicationGroup;
import com.amazonaws.services.elasticache.model.SecurityGroupMembership;
import com.amazonaws.services.elasticloadbalancing.model.Listener;
import com.amazonaws.services.elasticloadbalancing.model.ListenerDescription;
import com.amazonaws.services.elasticloadbalancing.model.LoadBalancerDescription;
import com.amazonaws.services.elasticloadbalancingv2.model.Action;
import com.amazonaws.services.elasticloadbalancingv2.model.AvailabilityZone;
import com.amazonaws.services.elasticloadbalancingv2.model.Certificate;
import com.amazonaws.services.elasticloadbalancingv2.model.DescribeListenersRequest;
import com.amazonaws.services.elasticloadbalancingv2.model.DescribeListenersResult;
import com.amazonaws.services.elasticloadbalancingv2.model.DescribeLoadBalancersRequest;
import com.amazonaws.services.elasticloadbalancingv2.model.DescribeRulesRequest;
import com.amazonaws.services.elasticloadbalancingv2.model.DescribeRulesResult;
import com.amazonaws.services.elasticloadbalancingv2.model.DescribeTagsRequest;
import com.amazonaws.services.elasticloadbalancingv2.model.DescribeTagsResult;
import com.amazonaws.services.elasticloadbalancingv2.model.DescribeTargetGroupAttributesRequest;
import com.amazonaws.services.elasticloadbalancingv2.model.DescribeTargetGroupAttributesResult;
import com.amazonaws.services.elasticloadbalancingv2.model.DescribeTargetGroupsRequest;
import com.amazonaws.services.elasticloadbalancingv2.model.DescribeTargetGroupsResult;
import com.amazonaws.services.elasticloadbalancingv2.model.DescribeTargetHealthRequest;
import com.amazonaws.services.elasticloadbalancingv2.model.DescribeTargetHealthResult;
import com.amazonaws.services.elasticloadbalancingv2.model.LoadBalancer;
import com.amazonaws.services.elasticloadbalancingv2.model.LoadBalancerAddress;
import com.amazonaws.services.elasticloadbalancingv2.model.Rule;
import com.amazonaws.services.elasticloadbalancingv2.model.RuleCondition;
import com.amazonaws.services.elasticloadbalancingv2.model.TagDescription;
import com.amazonaws.services.elasticloadbalancingv2.model.TargetDescription;
import com.amazonaws.services.elasticloadbalancingv2.model.TargetGroup;
import com.amazonaws.services.elasticloadbalancingv2.model.TargetGroupAttribute;
import com.amazonaws.services.elasticloadbalancingv2.model.TargetHealth;
import com.amazonaws.services.elasticloadbalancingv2.model.TargetHealthDescription;
import com.amazonaws.services.kms.model.AliasListEntry;
import com.amazonaws.services.kms.model.DescribeKeyRequest;
import com.amazonaws.services.kms.model.DescribeKeyResult;
import com.amazonaws.services.kms.model.KeyListEntry;
import com.amazonaws.services.kms.model.KeyMetadata;
import com.amazonaws.services.kms.model.ListAliasesResult;
import com.amazonaws.services.kms.model.ListKeysResult;
import com.amazonaws.services.rds.model.DBCluster;
import com.amazonaws.services.rds.model.DBClusterMember;
import com.amazonaws.services.rds.model.DBClusterOptionGroupStatus;
import com.amazonaws.services.rds.model.DBClusterRole;
import com.amazonaws.services.rds.model.DBInstance;
import com.amazonaws.services.rds.model.DBParameterGroupStatus;
import com.amazonaws.services.rds.model.DBSecurityGroupMembership;
import com.amazonaws.services.rds.model.DescribeDBClustersResult;
import com.amazonaws.services.rds.model.DescribeDBInstancesResult;
import com.amazonaws.services.rds.model.DomainMembership;
import com.amazonaws.services.rds.model.Endpoint;
import com.amazonaws.services.rds.model.VpcSecurityGroupMembership;
import com.amazonaws.services.s3.model.AbortIncompleteMultipartUpload;
import com.amazonaws.services.s3.model.AccessControlList;
import com.amazonaws.services.s3.model.AmazonS3Exception;
import com.amazonaws.services.s3.model.Bucket;
import com.amazonaws.services.s3.model.BucketAccelerateConfiguration;
import com.amazonaws.services.s3.model.BucketCrossOriginConfiguration;
import com.amazonaws.services.s3.model.BucketLifecycleConfiguration;
import com.amazonaws.services.s3.model.BucketLifecycleConfiguration.NoncurrentVersionTransition;
import com.amazonaws.services.s3.model.BucketLifecycleConfiguration.Transition;
import com.amazonaws.services.s3.model.BucketLoggingConfiguration;
import com.amazonaws.services.s3.model.BucketNotificationConfiguration;
import com.amazonaws.services.s3.model.BucketReplicationConfiguration;
import com.amazonaws.services.s3.model.BucketTaggingConfiguration;
import com.amazonaws.services.s3.model.BucketVersioningConfiguration;
import com.amazonaws.services.s3.model.BucketWebsiteConfiguration;
import com.amazonaws.services.s3.model.CORSRule;
import com.amazonaws.services.s3.model.CORSRule.AllowedMethods;
import com.amazonaws.services.s3.model.FilterRule;
import com.amazonaws.services.s3.model.GetBucketEncryptionResult;
import com.amazonaws.services.s3.model.Grant;
import com.amazonaws.services.s3.model.Grantee;
import com.amazonaws.services.s3.model.NotificationConfiguration;
import com.amazonaws.services.s3.model.RedirectRule;
import com.amazonaws.services.s3.model.ReplicationDestinationConfig;
import com.amazonaws.services.s3.model.ReplicationRule;
import com.amazonaws.services.s3.model.RoutingRule;
import com.amazonaws.services.s3.model.RoutingRuleCondition;
import com.amazonaws.services.s3.model.S3KeyFilter;
import com.amazonaws.services.s3.model.ServerSideEncryptionByDefault;
import com.amazonaws.services.s3.model.ServerSideEncryptionConfiguration;
import com.amazonaws.services.s3.model.ServerSideEncryptionRule;
import com.amazonaws.services.s3.model.SourceSelectionCriteria;
import com.amazonaws.services.s3.model.SseKmsEncryptedObjects;
import com.amazonaws.services.s3.model.TagSet;

import anthunt.aws.exporter.model.AmazonAccess;
import anthunt.poi.helper.XSSFHelper;

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
	private Regions region;
  
	private DateFormat formatDate = new SimpleDateFormat("yyyy-MM-dd HH:mm:ss");
	private DateFormat format = new SimpleDateFormat("yyyyMMddHHmmss");
  
	public AWSExporter(AmazonAccess amazonAccess, List<Regions> regions, boolean isCrossAccount, String crossAccountKey) {
		this.amazonAccess = amazonAccess;
		this.executeTime = format.format(new Date());
    
		System.out.println("AWS Exporter Start");
		for (Regions region : regions) {
			make(region, isCrossAccount, crossAccountKey);
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
		this.vpcSheet = this.workbook.createSheet("1.VPCs");
		this.vpcPeeringSheet = this.workbook.createSheet("2.VPCPeerings");
		this.subnetSheet = this.workbook.createSheet("3.Subnets");
		this.routeTableSheet = this.workbook.createSheet("4.Route Tables");
		this.internetGatewaySheet = this.workbook.createSheet("5.Internet Gateways");
		this.egressInternetGatewaySheet = this.workbook.createSheet("6.Egress Only Internet Gateways");
		this.natGatewaySheet = this.workbook.createSheet("7.NAT Gateways");
		this.customerGatewaySheet = this.workbook.createSheet("8.Customer Gateways");
		this.vpnGatewaySheet = this.workbook.createSheet("9.VPN Gateways");
		this.vpnConnectionSheet = this.workbook.createSheet("10.VPN Connections");
		this.securityGroupSheet = this.workbook.createSheet("11.Security Groups");
		this.ec2InstanceSheet = this.workbook.createSheet("12.EC2 Instances");
		this.ebsSheet = this.workbook.createSheet("13.EBS Volumes");
		this.classicElbSheet = this.workbook.createSheet("14.Classic ELBs");
		this.otherElbSheet = this.workbook.createSheet("15.Other ELBs");
		this.elastiCacheSheet = this.workbook.createSheet("16.ElastiCaches");
		this.rdsClusterSheet = this.workbook.createSheet("17.RDS-Clusters");
		this.rdsInstanceSheet = this.workbook.createSheet("18.RDS-Instances");
		this.kmsSheet = this.workbook.createSheet("19.KMS");
		this.acmSheet = this.workbook.createSheet("20.ACM");
		this.s3Sheet = this.workbook.createSheet("21.S3");
		this.directConnectSheet = this.workbook.createSheet("22.Direct Connects");
		this.directLocationSheet = this.workbook.createSheet("23.Direct Connect Locations");
		this.virtualGatewaySheet = this.workbook.createSheet("24.Virtual Gateways");
		this.virtualInterfaceSheet = this.workbook.createSheet("25.Virtual Interfaces");
		this.lagSheet = this.workbook.createSheet("26.Lags");
		this.directConnectGatewaySheet = this.workbook.createSheet("27.Direct Connect Gateways");
		this.directorySheet = this.workbook.createSheet("28.Directory Services");
		
		this.xssfHelper = new XSSFHelper(this.workbook);
	}
  
	private String getNameTagValue(List<Tag> tags) {
		String name = "";
		for (Tag tag : tags) {
			if ("Name".equals(tag.getKey())) {
				name = tag.getValue();
				break;
			}
		}
		return name;
	}
  
	private String getAllTagValue(List<Tag> tags) {
		StringBuffer tagValues = new StringBuffer();
		for (Tag tag : tags) {
			if(tagValues.length() > 0) {
				tagValues.append("\n");
			}
			tagValues.append(tag.getKey());
			tagValues.append("=");
			tagValues.append(tag.getValue());
		}
		return tagValues.toString();
	}
  
	private void make(Regions region, boolean isCrossAccountRole, String crossAccountKey) {
		this.region = region;
    
		String typeName = isCrossAccountRole ? crossAccountKey : this.amazonAccess.getAccessType().toUpperCase();
    
		this.fileName = ("[" + typeName + "][" + this.region.getName() + "] AWSExport-" + this.executeTime + ".xlsx");
		System.out.println("Start " + this.region.getName() + " Export. OutputFile [" + this.fileName + "]");
    
		initializeWorkbook();
		System.out.println("Excel Workbook 초기화 완료");
    
		this.amazonClients = new AmazonClients(this.region, this.amazonAccess, isCrossAccountRole, crossAccountKey);
		System.out.println("AWS Client 초기화 완료");
    
		makeSheetListHeader();
		System.out.println("Excel Workbook Sheet 생성 및 제목 줄 생성 완료 [" + this.workbook.getNumberOfSheets() + " 개 Sheet]");
		
		try { makeSubnet(); System.out.println("Subnets Export 완료"); } catch (Exception e) { System.out.println("Subnets Export 실패 - [" + e.getMessage() + "]"); e.printStackTrace(); }
		
		HashMap<String, String> vpcMainRouteTables = null;
		try { vpcMainRouteTables = makeRouteTable(); System.out.println("RouteTable Export 완료"); } catch (Exception e) { System.out.println("RouteTable Export 실패 - [" + e.getMessage() + "]"); e.printStackTrace(); }
		try { makeVPC(vpcMainRouteTables); System.out.println("VPC Export 완료"); } catch (Exception e) { System.out.println("VPC Export 실패 - [" + e.getMessage() + "]"); e.printStackTrace(); }
		try { makeVPCPeering(); System.out.println("VPC Export 완료"); } catch (Exception e) { System.out.println("VPC Export 실패 - [" + e.getMessage() + "]"); e.printStackTrace(); }
		try { makeInternetGateway(); System.out.println("InternetGateway Export 완료"); } catch (Exception e) { System.out.println("InternetGateway Export 실패 - [" + e.getMessage() + "]"); e.printStackTrace(); }
		try { makeEgressInternetGateway(); System.out.println("InternetGateway Export 완료"); } catch (Exception e) { System.out.println("InternetGateway Export 실패 - [" + e.getMessage() + "]"); e.printStackTrace(); }
		try { makeNATGateway(); System.out.println("EgressOnlyInternetGateway Export 완료"); } catch (Exception e) { System.out.println("EgressOnlyInternetGateway Export 실패 - [" + e.getMessage() + "]"); e.printStackTrace(); }
		try { makeCustomerGateway(); System.out.println("CustomerGateway Export 완료"); } catch (Exception e) { System.out.println("CustomerGateway Export 실패 - [" + e.getMessage() + "]"); e.printStackTrace(); }
		try { makeVPNGateway(); System.out.println("VPNGateway Export 완료"); } catch (Exception e) { System.out.println("VPNGateway Export 실패 - [" + e.getMessage() + "]"); e.printStackTrace(); }
		try { makeVPNConnection(); System.out.println("VPNConnection Export 완료"); } catch (Exception e) { System.out.println("VPNConnection Export 실패 - [" + e.getMessage() + "]"); e.printStackTrace(); }		
		try { makeSecurityGroup(); System.out.println("SecurityGroup Export 완료"); } catch (Exception e) { System.out.println("SecurityGroup Export 실패 - [" + e.getMessage() + "]"); e.printStackTrace(); }
		try { makeEC2Instance(); System.out.println("EC2Instances Export 완료"); } catch (Exception e) { System.out.println("EC2Instances Export 실패 - [" + e.getMessage() + "]"); e.printStackTrace(); }
		try { makeEBS(); System.out.println("EBS Export 완료"); } catch (Exception e) { System.out.println("EBS Export 실패 - [" + e.getMessage() + "]"); e.printStackTrace(); }
		try { makeClassicELB(); System.out.println("Classic ELB Export 완료"); } catch (Exception e) { System.out.println("Classic ELB Export 실패 - [" + e.getMessage() + "]"); e.printStackTrace(); }
		try { makeOtherELB(); System.out.println("Other ELB Export 완료"); } catch (Exception e) { System.out.println("Other ELB Export 실패 - [" + e.getMessage() + "]"); e.printStackTrace(); }
		try { makeElasticCache(); System.out.println("ElasticCache Export 완료"); } catch (Exception e) { System.out.println("ElasticCache Export 실패 - [" + e.getMessage() + "]"); e.printStackTrace(); }
		try { makeRDSCluster(); System.out.println("RDS Clusters Export 완료"); } catch (Exception e) { System.out.println("RDS Clusters Export 실패 - [" + e.getMessage() + "]"); e.printStackTrace(); }
		try { makeRDSInstance(); System.out.println("RDS Instances Export 완료"); } catch (Exception e) { System.out.println("RDS Instances Export 실패 - [" + e.getMessage() + "]"); e.printStackTrace(); }
		try { makeKMS(); System.out.println("KMS Export 완료"); } catch (Exception e) { System.out.println("KMS Export 실패 - [" + e.getMessage() + "]"); e.printStackTrace(); }
		try { makeACM(); System.out.println("ACM Export 완료"); } catch (Exception e) { System.out.println("ACM Export 실패 - [" + e.getMessage() + "]"); e.printStackTrace(); }
		try { makeS3(); System.out.println("S3 Export 완료"); } catch (Exception e) { System.out.println("S3 Export 실패 - [" + e.getMessage() + "]"); e.printStackTrace(); }
		
		try { makeDirectConnection(); System.out.println("DirectConnect Export 완료"); } catch (Exception e) { System.out.println("DirectConnect Export 실패 - [" + e.getMessage() + "]"); e.printStackTrace(); }
		try { makeDirectLocation(); System.out.println("DirectLocation Export 완료"); } catch (Exception e) { System.out.println("DirectLocation Export 실패 - [" + e.getMessage() + "]"); e.printStackTrace(); }
		try { makeVirtualGateway(); System.out.println("VirtualGateway Export 완료"); } catch (Exception e) { System.out.println("VirtualGateway Export 실패 - [" + e.getMessage() + "]"); e.printStackTrace(); }
		try { makeVirtualInterface(); System.out.println("VirtualInterface Export 완료"); } catch (Exception e) { System.out.println("VirtualInterface Export 실패 - [" + e.getMessage() + "]"); e.printStackTrace(); }
		try { makeLag(); System.out.println("Lag Export 완료"); } catch (Exception e) { System.out.println("Lag Export 실패 - [" + e.getMessage() + "]"); e.printStackTrace(); }
		try { makeDirectConnectGateway(); System.out.println("DirectConnect Gateway Export 완료"); } catch (Exception e) { System.out.println("DirectConnect Gateway Export 실패 - [" + e.getMessage() + "]"); e.printStackTrace(); }
		
		try { makeDirectoryService(); System.out.println("Directory Service Export 완료"); } catch (Exception e) { System.out.println("Directory Service Export 실패 - [" + e.getMessage() + "]"); e.printStackTrace(); }
		
		this.xssfHelper.setAutoSizeColumn();
	}
	
	private void makeDirectoryService() {
		XSSFRow row = null;
		DescribeDirectoriesResult describeDirectoriesResult = this.amazonClients.AwsDirectoryService.describeDirectories();
		List<DirectoryDescription> directoryDescriptions = describeDirectoriesResult.getDirectoryDescriptions();
		for (DirectoryDescription directoryDescription : directoryDescriptions) {
			row = this.xssfHelper.createRow(this.directorySheet, 1);
			this.xssfHelper.setLeftThinCell(row, Integer.toString(row.getRowNum() - 1));
			this.xssfHelper.setCell(row, directoryDescription.getDirectoryId());
			this.xssfHelper.setCell(row, directoryDescription.getType());
			this.xssfHelper.setCell(row, directoryDescription.getAlias());
			this.xssfHelper.setCell(row, directoryDescription.getName());
			this.xssfHelper.setCell(row, directoryDescription.getShortName());
			this.xssfHelper.setCell(row, directoryDescription.getSize());
			this.xssfHelper.setCell(row, directoryDescription.getAccessUrl());
			this.xssfHelper.setCell(row, directoryDescription.getDesiredNumberOfDomainControllers().toString());
			
			DirectoryVpcSettingsDescription directoryVpcSettingsDescription = directoryDescription.getVpcSettings();
			this.xssfHelper.setCell(row, directoryVpcSettingsDescription == null ? "" : directoryVpcSettingsDescription.getVpcId());
			this.xssfHelper.setCell(row, directoryVpcSettingsDescription == null ? "" : directoryVpcSettingsDescription.getSecurityGroupId());
			
			StringBuffer vpcAvas = new StringBuffer();
			if(directoryVpcSettingsDescription != null) {
			List<String> availabilityZones = directoryVpcSettingsDescription.getAvailabilityZones();
				for(String availabilityZone : availabilityZones) {
					if(vpcAvas.length() > 0) vpcAvas.append("\n");
					vpcAvas.append(availabilityZone);
				}
			}
			this.xssfHelper.setCell(row, vpcAvas.toString());
			
			StringBuffer vpcSubs = new StringBuffer();
			if(directoryVpcSettingsDescription != null) {
			List<String> subnetIds = directoryVpcSettingsDescription.getSubnetIds();
				for(String subnetId : subnetIds) {
					if(vpcSubs.length() > 0) vpcSubs.append("\n");
					vpcSubs.append(subnetId);
				}
			}
			this.xssfHelper.setCell(row, vpcSubs.toString());
			
			DirectoryConnectSettingsDescription directoryConnectSettingsDescription = directoryDescription.getConnectSettings();
			this.xssfHelper.setCell(row, directoryConnectSettingsDescription == null ? "" : directoryConnectSettingsDescription.getCustomerUserName());
			this.xssfHelper.setCell(row, directoryConnectSettingsDescription == null ? "" : directoryConnectSettingsDescription.getVpcId());
			this.xssfHelper.setCell(row, directoryConnectSettingsDescription == null ? "" : directoryConnectSettingsDescription.getSecurityGroupId());
			
			StringBuffer csAvas = new StringBuffer();
			if(directoryConnectSettingsDescription != null) {
				List<String> csavailabilityZones = directoryConnectSettingsDescription.getAvailabilityZones();
				for(String availabilityZone : csavailabilityZones) {
					if(csAvas.length() > 0) csAvas.append("\n");
					csAvas.append(availabilityZone);
				}
			}
			this.xssfHelper.setCell(row, csAvas.toString());
			
			StringBuffer cips = new StringBuffer();
			if(directoryConnectSettingsDescription != null) {
				List<String> connectIps = directoryConnectSettingsDescription.getConnectIps();
				for(String connectIp : connectIps) {
					if(cips.length() > 0) cips.append("\n");
					cips.append(connectIp);
				}
			}
			this.xssfHelper.setCell(row, cips.toString());
			
			StringBuffer csSubs = new StringBuffer();
			if(directoryConnectSettingsDescription != null) {
				List<String> csSubnetIds = directoryConnectSettingsDescription.getSubnetIds();
				for(String csSubnetId : csSubnetIds) {
					if(csSubs.length() > 0) csSubs.append("\n");
					csSubs.append(csSubnetId);
				}
			}			
			this.xssfHelper.setCell(row, csSubs.toString());
			
			this.xssfHelper.setCell(row, directoryConnectSettingsDescription == null ? "" : directoryDescription.getDescription());
			
			StringBuffer dias = new StringBuffer();
			if(directoryConnectSettingsDescription != null) {
				List<String> dnsIpAddrs = directoryDescription.getDnsIpAddrs();
				for(String dnsIpAddr : dnsIpAddrs) {
					if(dias.length() > 0) dias.append("\n");
					dias.append(dnsIpAddr);
				}
			}
			this.xssfHelper.setCell(row, dias.toString());
			this.xssfHelper.setCell(row, directoryConnectSettingsDescription == null ? "" : directoryDescription.getEdition());
			this.xssfHelper.setCell(row, directoryConnectSettingsDescription == null ? "" : directoryDescription.getSsoEnabled().toString());
			this.xssfHelper.setCell(row, directoryConnectSettingsDescription == null ? "" : directoryDescription.getStage());
			this.xssfHelper.setCell(row, directoryConnectSettingsDescription == null ? "" : directoryDescription.getStageLastUpdatedDateTime().toString());
			this.xssfHelper.setCell(row, directoryConnectSettingsDescription == null ? "" : directoryDescription.getStageReason());
			
			RadiusSettings radiusSettings = directoryConnectSettingsDescription == null ? null : directoryDescription.getRadiusSettings();
			this.xssfHelper.setCell(row, radiusSettings == null ? "" : radiusSettings.getAuthenticationProtocol());
			this.xssfHelper.setCell(row, radiusSettings == null ? "" : radiusSettings.getDisplayLabel());
			this.xssfHelper.setCell(row, radiusSettings == null ? "" : radiusSettings.getRadiusPort().toString());
			this.xssfHelper.setCell(row, radiusSettings == null ? "" : radiusSettings.getRadiusRetries().toString());
			this.xssfHelper.setCell(row, radiusSettings == null ? "" : radiusSettings.getRadiusTimeout().toString());
			this.xssfHelper.setCell(row, radiusSettings == null ? "" : radiusSettings.getSharedSecret());
			this.xssfHelper.setCell(row, radiusSettings == null ? "" : radiusSettings.getUseSameUsername().toString());
			
			StringBuffer rss = new StringBuffer();
			if(radiusSettings != null) {
				List<String> radiusServers = radiusSettings.getRadiusServers();
				for(String radiusServer : radiusServers) {
					if(rss.length() > 0) rss.append("\n");
					rss.append(radiusServer);
				}
			}
			this.xssfHelper.setCell(row, rss.toString());
			this.xssfHelper.setCell(row, directoryDescription.getRadiusStatus());
			
			StringBuffer dcrs = new StringBuffer();
			DescribeDomainControllersResult describeDomainControllersResult = this.amazonClients.AwsDirectoryService.describeDomainControllers(new DescribeDomainControllersRequest().withDirectoryId(directoryDescription.getDirectoryId()));
			List<DomainController> domainControllers = describeDomainControllersResult.getDomainControllers();
			for(DomainController domainController : domainControllers) {
				if(dcrs.length() > 0) dcrs.append("\n\n");
				dcrs.append("Domain Controller");
				dcrs.append("\nDomain Controller Id=");
				dcrs.append(domainController.getDomainControllerId());
				dcrs.append("\nDirectory Id=");
				dcrs.append(domainController.getDirectoryId());
				dcrs.append("\nVpc Id=");
				dcrs.append(domainController.getVpcId());
				dcrs.append("\nSubnet Id=");
				dcrs.append(domainController.getSubnetId());
				dcrs.append("\nAvailabilityZone=");
				dcrs.append(domainController.getAvailabilityZone());
				dcrs.append("\nDNS Ip Address=");
				dcrs.append(domainController.getDnsIpAddr());
				dcrs.append("\nLaunch Time=");
				dcrs.append(domainController.getLaunchTime().toString());
				dcrs.append("\nStatus=");
				dcrs.append(domainController.getStatus());
				dcrs.append("\nStatus Last Updated Date Time=");
				dcrs.append(domainController.getStatusLastUpdatedDateTime().toString());
				dcrs.append("\nStatus Reason=");
				dcrs.append(domainController.getStatusReason());
			}
			this.xssfHelper.setCell(row, dcrs.toString());
			
			StringBuffer trs = new StringBuffer();
			DescribeTrustsResult describeTrustsResult = this.amazonClients.AwsDirectoryService.describeTrusts(new DescribeTrustsRequest().withDirectoryId(directoryDescription.getDirectoryId()));
			List<Trust> trusts = describeTrustsResult.getTrusts();
			for(Trust trust : trusts) {
				if(trs.length() > 0) trs.append("\n\n");
				trs.append("Trusts");
				trs.append("\nDirectory Id=");
				trs.append(trust.getDirectoryId());
				trs.append("\nRemote Domain Name=");
				trs.append(trust.getRemoteDomainName());
				trs.append("\nId=");
				trs.append(trust.getTrustId());
				trs.append("\nType=");
				trs.append(trust.getTrustType());
				trs.append("\nDirection=");
				trs.append(trust.getTrustDirection());
				trs.append("\nState=");
				trs.append(trust.getTrustState());
				trs.append("\nState Reason=");
				trs.append(trust.getTrustStateReason());
				trs.append("\nCreated Date Time=");
				trs.append(trust.getCreatedDateTime().toString());
				trs.append("\nLast Updated Date Time=");
				trs.append(trust.getLastUpdatedDateTime().toString());
				trs.append("\nState Last Updated Date Time=");
				trs.append(trust.getStateLastUpdatedDateTime().toString());
			}
			
			this.xssfHelper.setCell(row, trs.toString());
			
			StringBuffer cfrs = new StringBuffer();
			try {
				DescribeConditionalForwardersResult describeConditionalForwardersResult = this.amazonClients.AwsDirectoryService.describeConditionalForwarders(new DescribeConditionalForwardersRequest().withDirectoryId(directoryDescription.getDirectoryId()));
				List<ConditionalForwarder> conditionalForwarders = describeConditionalForwardersResult.getConditionalForwarders();
				for(ConditionalForwarder conditionalForwarder : conditionalForwarders) {
					if(cfrs.length() > 0) cfrs.append("\n\n");
					cfrs.append("Conditional Forwarder");
					cfrs.append("\nRemote Domain Name=");
					cfrs.append(conditionalForwarder.getRemoteDomainName());
					cfrs.append("\nReplication Scope=");
					cfrs.append(conditionalForwarder.getReplicationScope());
					
					StringBuffer cdias = new StringBuffer();
					List<String> dnsIpAddrs = conditionalForwarder.getDnsIpAddrs();
					for(String dnsIpAddr : dnsIpAddrs) {
						if(cdias.length() > 0) cdias.append(", ");
						cdias.append(dnsIpAddr);
					}
					cfrs.append("\nDNS Ip Addresses=");
					cfrs.append(cdias.toString());
				}
			} catch(InvalidParameterException skip) {}
			this.xssfHelper.setCell(row, cfrs.toString());
			this.xssfHelper.setRightThinCell(row, directoryDescription.getLaunchTime().toString());
			
		}

	}
	
	private void makeDirectConnection() {
		XSSFRow row = null;
		DescribeConnectionsResult describeConnectionsResult = this.amazonClients.AmazonDirectConnect.describeConnections();
		List<Connection> connections = describeConnectionsResult.getConnections();
		for (Connection connection : connections) {
			row = this.xssfHelper.createRow(this.directConnectSheet, 1);
			this.xssfHelper.setLeftThinCell(row, Integer.toString(row.getRowNum() - 1));
			this.xssfHelper.setCell(row, connection.getRegion());
			this.xssfHelper.setCell(row, connection.getOwnerAccount());
			this.xssfHelper.setCell(row, connection.getConnectionId());
			this.xssfHelper.setCell(row, connection.getConnectionName());
			this.xssfHelper.setCell(row, connection.getConnectionState());
			this.xssfHelper.setCell(row, connection.getLocation());
			this.xssfHelper.setCell(row, connection.getAwsDevice());
			this.xssfHelper.setCell(row, connection.getVlan() == null ? "" : connection.getVlan().toString());
			this.xssfHelper.setCell(row, connection.getBandwidth());
			this.xssfHelper.setCell(row, connection.getPartnerName());
			this.xssfHelper.setCell(row, connection.getLagId());
			this.xssfHelper.setRightThinCell(row, connection.getLoaIssueTime().toString());
		}
		
	}
	
	private void makeDirectLocation() {
		XSSFRow row = null;
		DescribeLocationsResult describeLocationsResult = this.amazonClients.AmazonDirectConnect.describeLocations();
		List<Location> locations = describeLocationsResult.getLocations();
		for (Location location : locations) {
			row = this.xssfHelper.createRow(this.directLocationSheet, 1);
			this.xssfHelper.setLeftThinCell(row, Integer.toString(row.getRowNum() - 1));
			this.xssfHelper.setCell(row, this.region.getName());
			this.xssfHelper.setCell(row, location.getLocationCode());
			this.xssfHelper.setRightThinCell(row, location.getLocationName());
		}
		
	}
	
	private void makeVirtualGateway() {
		XSSFRow row = null;
		DescribeVirtualGatewaysResult describeVirtualGatewaysResult = this.amazonClients.AmazonDirectConnect.describeVirtualGateways();
		List<VirtualGateway> virtualGateways = describeVirtualGatewaysResult.getVirtualGateways();
		for (VirtualGateway virtualGateway : virtualGateways) {
			row = this.xssfHelper.createRow(this.virtualGatewaySheet, 1);
			this.xssfHelper.setLeftThinCell(row, Integer.toString(row.getRowNum() - 1));
			this.xssfHelper.setCell(row, this.region.getName());
			this.xssfHelper.setCell(row, virtualGateway.getVirtualGatewayId());
			this.xssfHelper.setRightThinCell(row, virtualGateway.getVirtualGatewayState());
		}
		
	}
	
	private void makeVirtualInterface() {
		XSSFRow row = null;
		DescribeVirtualInterfacesResult describeVirtualInterfacesResult = this.amazonClients.AmazonDirectConnect.describeVirtualInterfaces();
		List<VirtualInterface> virtualInterfaces = describeVirtualInterfacesResult.getVirtualInterfaces();
		for (VirtualInterface virtualInterface : virtualInterfaces) {
			row = this.xssfHelper.createRow(this.virtualInterfaceSheet, 1);
			this.xssfHelper.setLeftThinCell(row, Integer.toString(row.getRowNum() - 1));
			this.xssfHelper.setCell(row, this.region.getName());
			this.xssfHelper.setCell(row, virtualInterface.getVirtualGatewayId());
			this.xssfHelper.setCell(row, virtualInterface.getVirtualInterfaceId());
			this.xssfHelper.setCell(row, virtualInterface.getVirtualInterfaceName());
			this.xssfHelper.setCell(row, virtualInterface.getVirtualInterfaceState());
			this.xssfHelper.setCell(row, virtualInterface.getVirtualInterfaceType());
			this.xssfHelper.setCell(row, virtualInterface.getAsn().toString());
			this.xssfHelper.setCell(row, virtualInterface.getAddressFamily());
			this.xssfHelper.setCell(row, virtualInterface.getAmazonSideAsn().toString());
			this.xssfHelper.setCell(row, virtualInterface.getAmazonAddress());
			this.xssfHelper.setCell(row, virtualInterface.getAuthKey());
			
			StringBuffer bgps = new StringBuffer();
			List<BGPPeer> bgpPeers = virtualInterface.getBgpPeers();
			for(BGPPeer bgpPeer : bgpPeers) {
				if(bgps.length() > 0) bgps.append("\n\n");
				bgps.append("ASN=");
				bgps.append(bgpPeer.getAsn());
				bgps.append("\nAddress Family=");
				bgps.append(bgpPeer.getAddressFamily());
				bgps.append("\nAmazon Address=");
				bgps.append(bgpPeer.getAmazonAddress());
				bgps.append("\nCustomer Address=");
				bgps.append(bgpPeer.getCustomerAddress());
				bgps.append("\nAuth Key=");
				bgps.append(bgpPeer.getAuthKey());
				bgps.append("\nBGP Peer State=");
				bgps.append(bgpPeer.getBgpPeerState());
				bgps.append("\nBGP Status=");
				bgps.append(bgpPeer.getBgpStatus());
			}
			this.xssfHelper.setCell(row, bgps.toString());
			this.xssfHelper.setCell(row, virtualInterface.getConnectionId());
			this.xssfHelper.setCell(row, virtualInterface.getCustomerAddress());
			this.xssfHelper.setCell(row, virtualInterface.getCustomerRouterConfig());
			this.xssfHelper.setCell(row, virtualInterface.getDirectConnectGatewayId());
			this.xssfHelper.setCell(row, virtualInterface.getLocation());
			this.xssfHelper.setCell(row, virtualInterface.getOwnerAccount());
			
			StringBuffer rfps = new StringBuffer();
			List<RouteFilterPrefix> routeFilterPrefixs = virtualInterface.getRouteFilterPrefixes();
			for(RouteFilterPrefix routeFilterPrefix : routeFilterPrefixs) {
				if(rfps.length() > 0) rfps.append("\n");
				rfps.append(routeFilterPrefix.getCidr());
			}
			this.xssfHelper.setCell(row, rfps.toString());
			this.xssfHelper.setRightThinCell(row, virtualInterface.getVlan().toString());
		}
		
	}
	
	private void makeLag() {
		XSSFRow row = null;
		DescribeLagsResult describeLagsResult = this.amazonClients.AmazonDirectConnect.describeLags(new DescribeLagsRequest());
		List<Lag> lags = describeLagsResult.getLags();
		for (Lag lag : lags) {
			row = this.xssfHelper.createRow(this.lagSheet, 1);
			this.xssfHelper.setLeftThinCell(row, Integer.toString(row.getRowNum() - 1));
			this.xssfHelper.setCell(row, lag.getRegion());
			this.xssfHelper.setCell(row, lag.getAllowsHostedConnections().toString());
			this.xssfHelper.setCell(row, lag.getAwsDevice());
			
			StringBuffer cons = new StringBuffer();
			List<Connection> connections = lag.getConnections();
			for(Connection connection : connections) {
				if(cons.length() > 0) cons.append("\n");
				cons.append(connection.getConnectionId());
				cons.append("|");
				cons.append(connection.getConnectionName());
				cons.append(" (");
				cons.append(connection.getConnectionState());
				cons.append(")");
			}
			this.xssfHelper.setCell(row, cons.toString());
			this.xssfHelper.setCell(row, lag.getConnectionsBandwidth());
			this.xssfHelper.setCell(row, lag.getLagId());
			this.xssfHelper.setCell(row, lag.getLagName());
			this.xssfHelper.setCell(row, lag.getLagState());
			this.xssfHelper.setCell(row, lag.getMinimumLinks().toString());
			this.xssfHelper.setCell(row, lag.getNumberOfConnections().toString());
			this.xssfHelper.setRightThinCell(row, lag.getOwnerAccount());			
		}
		
	}
	
	private void makeDirectConnectGateway() {
		XSSFRow row = null;
		DescribeDirectConnectGatewaysResult describeDirectConnectGatewaysResult = this.amazonClients.AmazonDirectConnect.describeDirectConnectGateways(new DescribeDirectConnectGatewaysRequest());
		List<DirectConnectGateway> directConnectGateways = describeDirectConnectGatewaysResult.getDirectConnectGateways();
		for (DirectConnectGateway directConnectGateway : directConnectGateways) {
			row = this.xssfHelper.createRow(this.directConnectGatewaySheet, 1);
			this.xssfHelper.setLeftThinCell(row, Integer.toString(row.getRowNum() - 1));
			this.xssfHelper.setCell(row, this.region.getName());
			this.xssfHelper.setCell(row, directConnectGateway.getDirectConnectGatewayId());
			this.xssfHelper.setCell(row, directConnectGateway.getDirectConnectGatewayName());
			this.xssfHelper.setCell(row, directConnectGateway.getAmazonSideAsn().toString());
			this.xssfHelper.setCell(row, directConnectGateway.getDirectConnectGatewayState());
			this.xssfHelper.setCell(row, directConnectGateway.getOwnerAccount());
			this.xssfHelper.setRightThinCell(row, directConnectGateway.getStateChangeError());
		}
		
		
	}
	
	private void makeEgressInternetGateway() {
		XSSFRow row = null;
		
		DescribeEgressOnlyInternetGatewaysResult describeEgressOnlyInternetGatewaysResult = this.amazonClients.AmazonEC2.describeEgressOnlyInternetGateways(new DescribeEgressOnlyInternetGatewaysRequest());
		List<EgressOnlyInternetGateway> egressOnlyInternetGateways = describeEgressOnlyInternetGatewaysResult.getEgressOnlyInternetGateways();
		for (EgressOnlyInternetGateway egressOnlyInternetGateway : egressOnlyInternetGateways) {

			row = this.xssfHelper.createRow(this.egressInternetGatewaySheet, 1);
			this.xssfHelper.setLeftThinCell(row, Integer.toString(row.getRowNum() - 1));
			this.xssfHelper.setCell(row, this.region.getName());
			this.xssfHelper.setCell(row, egressOnlyInternetGateway.getEgressOnlyInternetGatewayId());
			
			StringBuffer iga = new StringBuffer();
			List<InternetGatewayAttachment> internetGatewayAttachments = egressOnlyInternetGateway.getAttachments();
			for(InternetGatewayAttachment internetGatewayAttachment : internetGatewayAttachments) {
				if(iga.length() > 0) iga.append("\n");
				iga.append(internetGatewayAttachment.getVpcId());
				iga.append("(");
				iga.append(internetGatewayAttachment.getState());
				iga.append(")");
			}
			this.xssfHelper.setRightThinCell(row, iga.toString());
			
		}
		
	}
	
	private void makeNATGateway() {
		XSSFRow row = null;
		
		DescribeNatGatewaysResult describeNatGatewaysResult = this.amazonClients.AmazonEC2.describeNatGateways(new DescribeNatGatewaysRequest());
		List<NatGateway> natGateways = describeNatGatewaysResult.getNatGateways();
		for(NatGateway natGateway : natGateways) {
			row = this.xssfHelper.createRow(this.natGatewaySheet, 1);
			this.xssfHelper.setLeftThinCell(row, Integer.toString(row.getRowNum() - 1));
			this.xssfHelper.setCell(row, this.region.getName());
			this.xssfHelper.setCell(row, natGateway.getNatGatewayId());
			
			StringBuffer ngas = new StringBuffer();
			List<NatGatewayAddress> natGatewayAddresses = natGateway.getNatGatewayAddresses();
			for(NatGatewayAddress natGatewayAddress : natGatewayAddresses) {
				if(ngas.length() > 0) ngas.append("\n\n");
				ngas.append("\nNetworkInterfaceId=");
				ngas.append(natGatewayAddress.getNetworkInterfaceId());
				ngas.append("\nPrivate Ip=");
				ngas.append(natGatewayAddress.getPrivateIp());
				ngas.append("\nPublic Ip=");
				ngas.append(natGatewayAddress.getPublicIp());
				ngas.append("\nAllocation Id=");
				ngas.append(natGatewayAddress.getAllocationId());
			}
			this.xssfHelper.setCell(row, ngas.toString());
			this.xssfHelper.setCell(row, natGateway.getVpcId());
			this.xssfHelper.setCell(row, natGateway.getSubnetId());
			this.xssfHelper.setCell(row, natGateway.getFailureCode());
			this.xssfHelper.setCell(row, natGateway.getFailureMessage());
			ProvisionedBandwidth provisionedBandwidth = natGateway.getProvisionedBandwidth();
			this.xssfHelper.setCell(row, provisionedBandwidth == null ? "" : provisionedBandwidth.getProvisioned());
			this.xssfHelper.setCell(row, provisionedBandwidth == null ? "" : provisionedBandwidth.getProvisionTime().toString());
			this.xssfHelper.setCell(row, provisionedBandwidth == null ? "" : provisionedBandwidth.getRequested());
			this.xssfHelper.setCell(row, provisionedBandwidth == null ? "" : provisionedBandwidth.getRequestTime().toString());
			this.xssfHelper.setCell(row, provisionedBandwidth == null ? "" : provisionedBandwidth.getStatus());
			this.xssfHelper.setCell(row, natGateway.getState());
			this.xssfHelper.setCell(row, natGateway.getCreateTime().toString());
			this.xssfHelper.setCell(row, natGateway.getDeleteTime() == null ? "" : natGateway.getDeleteTime().toString());
			this.xssfHelper.setRightThinCell(row, this.getAllTagValue(natGateway.getTags()));
		}
		
	}
	
	private void makeCustomerGateway() {
		XSSFRow row = null;
		
		DescribeCustomerGatewaysResult describeCustomerGatewaysResult = this.amazonClients.AmazonEC2.describeCustomerGateways();
		List<CustomerGateway> customerGateways = describeCustomerGatewaysResult.getCustomerGateways();
		for(CustomerGateway customerGateway : customerGateways) {
			row = this.xssfHelper.createRow(this.customerGatewaySheet, 1);
			this.xssfHelper.setLeftThinCell(row, Integer.toString(row.getRowNum() - 1));
			this.xssfHelper.setCell(row, this.region.getName());
			this.xssfHelper.setCell(row, customerGateway.getCustomerGatewayId());
			this.xssfHelper.setCell(row, customerGateway.getType());
			this.xssfHelper.setCell(row, customerGateway.getIpAddress());
			this.xssfHelper.setCell(row, customerGateway.getBgpAsn());
			this.xssfHelper.setCell(row, customerGateway.getState());
			this.xssfHelper.setRightThinCell(row, this.getAllTagValue(customerGateway.getTags()));
		}
		
	}
	
	private void makeVPNGateway() {
		XSSFRow row = null;
		
		DescribeVpnGatewaysResult describeVpnGatewaysResult = this.amazonClients.AmazonEC2.describeVpnGateways();
		List<VpnGateway> vpnGateways = describeVpnGatewaysResult.getVpnGateways();
		for(VpnGateway vpnGateway : vpnGateways) {
			row = this.xssfHelper.createRow(this.vpnGatewaySheet, 1);
			this.xssfHelper.setLeftThinCell(row, Integer.toString(row.getRowNum() - 1));
			this.xssfHelper.setCell(row, this.region.getName());
			this.xssfHelper.setCell(row, vpnGateway.getVpnGatewayId());
			this.xssfHelper.setCell(row, vpnGateway.getAvailabilityZone());
			this.xssfHelper.setCell(row, vpnGateway.getType());
			this.xssfHelper.setCell(row, vpnGateway.getAmazonSideAsn().toString());
			this.xssfHelper.setCell(row, vpnGateway.getState());
			
			StringBuffer vas = new StringBuffer();
			List<VpcAttachment> vpcAttachments = vpnGateway.getVpcAttachments();
			for(VpcAttachment vpcAttachment : vpcAttachments) {
				if(vas.length() > 0) vas.append("\n");
				vas.append(vpcAttachment.getVpcId());
				vas.append(" (");
				vas.append(vpcAttachment.getState());
				vas.append(")");
			}
			this.xssfHelper.setCell(row, vas.toString());
			this.xssfHelper.setRightThinCell(row, this.getAllTagValue(vpnGateway.getTags()));
		}
		
	}
	
	private void makeVPNConnection() {
		XSSFRow row = null;
		
		DescribeVpnConnectionsResult describeVpnConnectionsResult = this.amazonClients.AmazonEC2.describeVpnConnections();
		List<VpnConnection> vpnConnections = describeVpnConnectionsResult.getVpnConnections();
		for(VpnConnection vpnConnection : vpnConnections) {
			row = this.xssfHelper.createRow(this.vpnConnectionSheet, 1);
			this.xssfHelper.setLeftThinCell(row, Integer.toString(row.getRowNum() - 1));
			this.xssfHelper.setCell(row, this.region.getName());
			this.xssfHelper.setCell(row, vpnConnection.getVpnConnectionId());
			this.xssfHelper.setCell(row, vpnConnection.getVpnGatewayId());
			this.xssfHelper.setCell(row, vpnConnection.getCustomerGatewayId());
			this.xssfHelper.setCell(row, vpnConnection.getCustomerGatewayConfiguration());
			
			StringBuffer vgts = new StringBuffer();
			List<VgwTelemetry> vgwTelemetries = vpnConnection.getVgwTelemetry();
			for(VgwTelemetry vgwTelemetry : vgwTelemetries) {
				if(vgts.length() > 0) vgts.append("\n\n");
				vgts.append("\nOutside IpAddress=");
				vgts.append(vgwTelemetry.getOutsideIpAddress());
				vgts.append("\nAccepted Route Count=");
				vgts.append(vgwTelemetry.getAcceptedRouteCount());
				vgts.append("\nLast Status Change=");
				vgts.append(vgwTelemetry.getLastStatusChange().toString());
				vgts.append("\nStatus=");
				vgts.append(vgwTelemetry.getStatus());
				vgts.append("\nStatus Message=");
				vgts.append(vgwTelemetry.getStatusMessage());
			}
			this.xssfHelper.setCell(row, vgts.toString());
			this.xssfHelper.setCell(row, vpnConnection.getType());
			this.xssfHelper.setCell(row, vpnConnection.getCategory());
			this.xssfHelper.setCell(row, vpnConnection.getOptions().getStaticRoutesOnly().toString());
			
			StringBuffer vsrs = new StringBuffer();
			List<VpnStaticRoute> vpnStaticRoutes = vpnConnection.getRoutes();
			for(VpnStaticRoute vpnStaticRoute : vpnStaticRoutes) {
				if(vsrs.length() > 0) vsrs.append("\n\n");
				vsrs.append("\nSource=");
				vsrs.append(vpnStaticRoute.getSource());
				vsrs.append("\nDestination=");
				vsrs.append(vpnStaticRoute.getDestinationCidrBlock());
				vsrs.append("\nState=");
				vsrs.append(vpnStaticRoute.getState());
			}
			this.xssfHelper.setCell(row, vsrs.toString());
			this.xssfHelper.setCell(row, vpnConnection.getState());
			this.xssfHelper.setRightThinCell(row, this.getAllTagValue(vpnConnection.getTags()));
		}
		
	}
	
	
	private void makeVPCPeering() {
	
		XSSFRow row = null;
		
		DescribeVpcPeeringConnectionsResult describeVpcPeeringConnectionsResult = this.amazonClients.AmazonEC2.describeVpcPeeringConnections();
		List<VpcPeeringConnection> vpcPeeringConnections = describeVpcPeeringConnectionsResult.getVpcPeeringConnections();
		for (VpcPeeringConnection vpcPeeringConnection : vpcPeeringConnections) {
			
			row = this.xssfHelper.createRow(this.vpcPeeringSheet, 1);
			this.xssfHelper.setLeftThinCell(row, Integer.toString(row.getRowNum() - 1));
			this.xssfHelper.setCell(row, this.region.getName());
			
			this.xssfHelper.setCell(row, vpcPeeringConnection.getVpcPeeringConnectionId());
			this.getVpcPeeringVPCInfo(row, vpcPeeringConnection.getRequesterVpcInfo());						
			this.getVpcPeeringVPCInfo(row, vpcPeeringConnection.getAccepterVpcInfo());
			this.xssfHelper.setCell(row, vpcPeeringConnection.getExpirationTime() == null ? "" : vpcPeeringConnection.getExpirationTime().toString());
			VpcPeeringConnectionStateReason vpcPeeringConnectionStateReason = vpcPeeringConnection.getStatus();
			this.xssfHelper.setCell(row, vpcPeeringConnectionStateReason == null ? "" : vpcPeeringConnectionStateReason.getCode());
			this.xssfHelper.setCell(row, vpcPeeringConnectionStateReason == null ? "" : vpcPeeringConnectionStateReason.getMessage());
			this.xssfHelper.setCell(row, this.getAllTagValue(vpcPeeringConnection.getTags()));
		}
	}
	
	private void getVpcPeeringVPCInfo(XSSFRow row, VpcPeeringConnectionVpcInfo vpcPeeringConnectionVpcInfo) {
		this.xssfHelper.setCell(row, vpcPeeringConnectionVpcInfo.getVpcId());
		this.xssfHelper.setCell(row, vpcPeeringConnectionVpcInfo.getRegion());
		this.xssfHelper.setCell(row, vpcPeeringConnectionVpcInfo.getCidrBlock());
		
		StringBuffer cidrs = new StringBuffer();
		List<CidrBlock> cidrBlocks = vpcPeeringConnectionVpcInfo.getCidrBlockSet();
		for(CidrBlock cidrBlock : cidrBlocks) {
			if(cidrs.length() > 0) cidrs.append("\n");
			cidrs.append(cidrBlock.getCidrBlock());
		}
		this.xssfHelper.setCell(row, cidrs.toString());
		
		StringBuffer ipv6s = new StringBuffer();
		List<Ipv6CidrBlock> ipv6CidrBlocks = vpcPeeringConnectionVpcInfo.getIpv6CidrBlockSet();
		for(Ipv6CidrBlock ipv6CidrBlock : ipv6CidrBlocks) {
			if(ipv6s.length() > 0) ipv6s.append("\n");
			ipv6s.append(ipv6CidrBlock.getIpv6CidrBlock());
		}
		this.xssfHelper.setCell(row, ipv6s.toString());
		this.xssfHelper.setCell(row, vpcPeeringConnectionVpcInfo.getOwnerId());
		VpcPeeringConnectionOptionsDescription vpcPeeringConnectionOptionsDescription = vpcPeeringConnectionVpcInfo.getPeeringOptions();
		this.xssfHelper.setCell(row, vpcPeeringConnectionOptionsDescription == null ? "" : vpcPeeringConnectionOptionsDescription.getAllowDnsResolutionFromRemoteVpc() == null ? "" : vpcPeeringConnectionOptionsDescription.getAllowDnsResolutionFromRemoteVpc().toString());
		this.xssfHelper.setCell(row, vpcPeeringConnectionOptionsDescription == null ? "" : vpcPeeringConnectionOptionsDescription.getAllowEgressFromLocalClassicLinkToRemoteVpc() == null ? "" : vpcPeeringConnectionOptionsDescription.getAllowEgressFromLocalClassicLinkToRemoteVpc().toString());
		this.xssfHelper.setRightThinCell(row, vpcPeeringConnectionOptionsDescription == null ? "" : vpcPeeringConnectionOptionsDescription.getAllowEgressFromLocalVpcToRemoteClassicLink() == null ? "" : vpcPeeringConnectionOptionsDescription.getAllowEgressFromLocalVpcToRemoteClassicLink().toString());
	}
	
	private void makeVPC(HashMap<String, String> vpcMainRouteTables) {
	    XSSFRow row = null;
	    List<Vpc> vpcs = this.amazonClients.AmazonEC2.describeVpcs().getVpcs();
	    for (int i = 0; i < vpcs.size(); i++)
	    {
			Vpc vpc = (Vpc)vpcs.get(i);
			  
			row = this.xssfHelper.createRow(this.vpcSheet, 1);
			this.xssfHelper.setLeftThinCell(row, Integer.toString(i + 1));
			this.xssfHelper.setCell(row, this.region.getName());
			  
			  
			this.xssfHelper.setCell(row, getNameTagValue(vpc.getTags()));
			this.xssfHelper.setCell(row, vpc.getVpcId());
			this.xssfHelper.setCell(row, vpc.getCidrBlock());
			this.xssfHelper.setCell(row, vpc.getDhcpOptionsId());
			this.xssfHelper.setCell(row, vpc.getInstanceTenancy());
			
			if (vpcMainRouteTables != null) {
			    this.xssfHelper.setCell(row, (String)vpcMainRouteTables.get(vpc.getVpcId()));
			} else {
			    this.xssfHelper.setCell(row, "");
			}
			
			List<String> vpcList = new ArrayList<String>();
			vpcList.add(vpc.getVpcId());
			NetworkAcl networkAcl = (NetworkAcl)this.amazonClients.AmazonEC2.describeNetworkAcls(new DescribeNetworkAclsRequest().withFilters(new Filter[] { new Filter("vpc-id", vpcList) })).getNetworkAcls().get(0);
			this.xssfHelper.setCell(row, networkAcl.getNetworkAclId() + " | " + getNameTagValue(networkAcl.getTags()));
			
			this.xssfHelper.setCell(row, Boolean.toString(vpc.getIsDefault()));
			this.xssfHelper.setCell(row, vpc.getState());
			 
			StringBuffer vpcCidr = new StringBuffer();
			List<VpcCidrBlockAssociation> vpcCidrBlockAssociations = vpc.getCidrBlockAssociationSet();
			for(VpcCidrBlockAssociation vpcCidrBlockAssociation : vpcCidrBlockAssociations) {
				  vpcCidr.append("AssociationId : " + vpcCidrBlockAssociation.getAssociationId());
				  vpcCidr.append("\n");
				  vpcCidr.append("CIDR Block : " + vpcCidrBlockAssociation.getCidrBlock());
				  VpcCidrBlockState vpcCidrBlockState = vpcCidrBlockAssociation.getCidrBlockState();
				  if(vpcCidrBlockState != null) {
					  vpcCidr.append("\n");
					  vpcCidr.append("State : " + vpcCidrBlockState.getState());
					  vpcCidr.append("\n");
					  vpcCidr.append("Status Message : " + vpcCidrBlockState.getStatusMessage());
				  }
			}
			this.xssfHelper.setCell(row, vpcCidr.toString());
			  
			StringBuffer vpcIpv6Cidr = new StringBuffer();
			List<VpcIpv6CidrBlockAssociation> vpcIpv6CidrBlockAssociations = vpc.getIpv6CidrBlockAssociationSet();
			for(VpcIpv6CidrBlockAssociation vpcIpv6CidrBlockAssociation : vpcIpv6CidrBlockAssociations) {
				vpcIpv6Cidr.append("AssociationId : " + vpcIpv6CidrBlockAssociation.getAssociationId());
				vpcIpv6Cidr.append("\n");
				vpcIpv6Cidr.append("CIDR Block : " + vpcIpv6CidrBlockAssociation.getIpv6CidrBlock());
				VpcCidrBlockState vpcCidrBlockState = vpcIpv6CidrBlockAssociation.getIpv6CidrBlockState();
				if(vpcCidrBlockState != null) {
					vpcIpv6Cidr.append("\n");
					vpcIpv6Cidr.append("State : " + vpcCidrBlockState.getState());
					vpcIpv6Cidr.append("\n");
					vpcIpv6Cidr.append("Status Message : " + vpcCidrBlockState.getStatusMessage());
				}
			}
			this.xssfHelper.setCell(row, vpcIpv6Cidr.toString());
			  
			this.xssfHelper.setRightThinCell(row, this.getAllTagValue(vpc.getTags()));
	      
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
		
		List<Bucket> buckets = this.amazonClients.AmazonS3.listBuckets();
		for (Bucket bucket : buckets) {
			
			row = this.xssfHelper.createRow(this.s3Sheet, 1);			
			this.xssfHelper.setLeftThinCell(row, Integer.toString(row.getRowNum() - 1));
			this.xssfHelper.setCell(row, bucket.getOwner().getDisplayName() + "(" + bucket.getOwner().getId() + ")");
			this.xssfHelper.setCell(row, bucket.getName());
			this.xssfHelper.setCell(row, bucket.getCreationDate().toString());
			
			try {
				this.xssfHelper.setCell(row, this.amazonClients.AmazonS3.getBucketLocation(bucket.getName()));

			} catch(AmazonS3Exception skip) {		
				this.xssfHelper.setCell(row, "");	
			} catch(Exception e) {
				this.xssfHelper.setCell(row, "");
				e.printStackTrace();
			}
			
			try {
				StringBuffer bets = new StringBuffer();
				GetBucketEncryptionResult getBucketEncryptionResult = this.amazonClients.AmazonS3.getBucketEncryption(bucket.getName());
				ServerSideEncryptionConfiguration serverSideEncryptionConfiguration = getBucketEncryptionResult.getServerSideEncryptionConfiguration();
				List<ServerSideEncryptionRule> serverSideEncryptionRules = serverSideEncryptionConfiguration.getRules();
				for(ServerSideEncryptionRule serverSideEncryptionRule : serverSideEncryptionRules) {
					ServerSideEncryptionByDefault serverSideEncryptionByDefault = serverSideEncryptionRule.getApplyServerSideEncryptionByDefault();
					
					if(bets.length() > 0) bets.append("\n");
					bets.append(serverSideEncryptionByDefault.getKMSMasterKeyID());
					bets.append("|");
					bets.append(serverSideEncryptionByDefault.getSSEAlgorithm());
				}
				this.xssfHelper.setCell(row,  bets.toString());
			} catch(AmazonS3Exception skip) {		
				this.xssfHelper.setCell(row, "");		
			} catch(Exception e) {	
				this.xssfHelper.setCell(row, "");
				e.printStackTrace();
			}
			
			try {
				BucketAccelerateConfiguration bucketAccelerateConfiguration = this.amazonClients.AmazonS3.getBucketAccelerateConfiguration(bucket.getName());
				this.xssfHelper.setCell(row,  Boolean.toString(bucketAccelerateConfiguration.isAccelerateEnabled()));
				this.xssfHelper.setCell(row,  bucketAccelerateConfiguration.getStatus());
			} catch(AmazonS3Exception skip) {
				this.xssfHelper.setCell(row, "");
				this.xssfHelper.setCell(row, "");
			} catch(Exception e) {
				this.xssfHelper.setCell(row, "");
				this.xssfHelper.setCell(row, "");
				e.printStackTrace();
			}
			
			try {
				StringBuffer crrs = new StringBuffer();
				BucketCrossOriginConfiguration bucketCrossOriginConfiguration = this.amazonClients.AmazonS3.getBucketCrossOriginConfiguration(bucket.getName());
				
				if(bucketCrossOriginConfiguration != null) {
					List<CORSRule> corsRules = bucketCrossOriginConfiguration.getRules();
					for(CORSRule corsRule : corsRules) {
						
						if(crrs.length() > 0) crrs.append("\n");
						crrs.append("Id=");
						crrs.append(corsRule.getId());
						crrs.append(", MaxAgeSeconds=");
						crrs.append(Integer.toString(corsRule.getMaxAgeSeconds()));
						
						crrs.append(", AllowedHeaders=");
						List<String> allowedHeaders = corsRule.getAllowedHeaders();
						for(String allowedHeader : allowedHeaders) {
							crrs.append(allowedHeader);
							crrs.append("|");
						}
						crrs.append(", AllowedMethods=");
						List<AllowedMethods> allowedMethods = corsRule.getAllowedMethods();
						for(AllowedMethods allowedMethod : allowedMethods) {
							crrs.append(allowedMethod.name());
							crrs.append("|");
						}
						crrs.append(", AllowedOrigins=");
						List<String> allowedOrigins = corsRule.getAllowedOrigins();
						for(String allowedOrigin : allowedOrigins) {
							crrs.append(allowedOrigin);
							crrs.append("|");
						}
						crrs.append(", ExposedHeaders=");
						List<String> exposedHeaders = corsRule.getExposedHeaders();
						for(String exposedHeader : exposedHeaders) {
							crrs.append(exposedHeader);
							crrs.append("|");
						}
						
					}
				}
				
				this.xssfHelper.setCell(row, crrs.toString());

			} catch(AmazonS3Exception skip) {
				this.xssfHelper.setCell(row, "");
			} catch(Exception e) {
				this.xssfHelper.setCell(row, "");
				e.printStackTrace();
			}
			
			try {
				this.xssfHelper.setCell(row, this.amazonClients.AmazonS3.getBucketPolicy(bucket.getName()).getPolicyText());
			} catch(AmazonS3Exception skip) {
				this.xssfHelper.setCell(row, "");
			} catch(Exception e) {
				this.xssfHelper.setCell(row, "");
				e.printStackTrace();
			}
			
			try {
				BucketLoggingConfiguration bucketLoggingConfiguration = this.amazonClients.AmazonS3.getBucketLoggingConfiguration(bucket.getName());
				this.xssfHelper.setCell(row, Boolean.toString(bucketLoggingConfiguration.isLoggingEnabled()));
				this.xssfHelper.setCell(row, bucketLoggingConfiguration.getDestinationBucketName());
				this.xssfHelper.setCell(row, bucketLoggingConfiguration.getLogFilePrefix());
			} catch(AmazonS3Exception skip) {
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
				BucketVersioningConfiguration bucketVersioningConfiguration = this.amazonClients.AmazonS3.getBucketVersioningConfiguration(bucket.getName());
				if(bucketVersioningConfiguration != null) {
					this.xssfHelper.setCell(row, bucketVersioningConfiguration.isMfaDeleteEnabled() == null ? "" : bucketVersioningConfiguration.isMfaDeleteEnabled().toString());
					this.xssfHelper.setCell(row, bucketVersioningConfiguration.getStatus());
				} else {
					this.xssfHelper.setCell(row, "");
					this.xssfHelper.setCell(row, "");
				}
			} catch(AmazonS3Exception skip) {
				this.xssfHelper.setCell(row, "");
				this.xssfHelper.setCell(row, "");
			} catch(Exception e) {
				this.xssfHelper.setCell(row, "");
				this.xssfHelper.setCell(row, "");
				e.printStackTrace();
			}
			
			try {
				StringBuffer btcs = new StringBuffer();
				BucketTaggingConfiguration bucketTaggingConfiguration = this.amazonClients.AmazonS3.getBucketTaggingConfiguration(bucket.getName());
				
				if(bucketTaggingConfiguration != null) {
					List<TagSet> tagSets = bucketTaggingConfiguration.getAllTagSets();
					for(TagSet tagSet : tagSets) {
						Map<String, String> allTagsMap = tagSet.getAllTags();
						Set<String> keys = allTagsMap.keySet();
						Iterator<String> iKeys = keys.iterator();
						while(iKeys.hasNext()) {
							if(btcs.length() > 0) btcs.append("\n");
							String key = iKeys.next();
							String value = allTagsMap.get(key);
							btcs.append(key);
							btcs.append("=");
							btcs.append(value);
						}
					}
				}
				this.xssfHelper.setCell(row, btcs.toString());
				
				StringBuffer tss = new StringBuffer();
				if(bucketTaggingConfiguration != null) {
					TagSet tagSet = bucketTaggingConfiguration.getTagSet();
					Map<String, String> allTagsMap = tagSet.getAllTags();
					Set<String> keys = allTagsMap.keySet();
					Iterator<String> iKeys = keys.iterator();
					while(iKeys.hasNext()) {
						if(tss.length() > 0) tss.append("\n");
						String key = iKeys.next();
						String value = allTagsMap.get(key);
						tss.append(key);
						tss.append("=");
						tss.append(value);
					}
				}
				this.xssfHelper.setCell(row, tss.toString());
			} catch(AmazonS3Exception skip) {
				this.xssfHelper.setCell(row, "");
				this.xssfHelper.setCell(row, "");
			} catch(Exception e) {
				this.xssfHelper.setCell(row, "");
				this.xssfHelper.setCell(row, "");
				e.printStackTrace();
			}
			
			try {
				BucketReplicationConfiguration bucketReplicationConfiguration = this.amazonClients.AmazonS3.getBucketReplicationConfiguration(bucket.getName());
				this.xssfHelper.setCell(row, bucketReplicationConfiguration.getRoleARN());
				
				StringBuffer rrs = new StringBuffer();
				Map<String, ReplicationRule> repMap = bucketReplicationConfiguration.getRules();
				Set<String> repkeys = repMap.keySet();
				Iterator<String> irepKeys = repkeys.iterator();
				while(irepKeys.hasNext()) {
					if(rrs.length() > 0) rrs.append("\n\n");
					String key = irepKeys.next();
					
					rrs.append(key);
					
					ReplicationRule replicationRule = repMap.get(key);
					ReplicationDestinationConfig replicationDestinationConfig = replicationRule.getDestinationConfig();
					rrs.append("\nAccessControlTranslation=");
					rrs.append(replicationDestinationConfig.getAccessControlTranslation());
					rrs.append("\nAccount=");
					rrs.append(replicationDestinationConfig.getAccount());
					rrs.append("\nBucketARN=");
					rrs.append(replicationDestinationConfig.getBucketARN());
					
					rrs.append("\nReplicaKmsKeyID=");
					rrs.append(replicationDestinationConfig.getEncryptionConfiguration().getReplicaKmsKeyID());
					
					rrs.append("\nStorageClass=");
					rrs.append(replicationDestinationConfig.getStorageClass());
					rrs.append("\nPrefix=");
					rrs.append(replicationRule.getPrefix());
					SourceSelectionCriteria sourceSelectionCriteria = replicationRule.getSourceSelectionCriteria();
					SseKmsEncryptedObjects sseKmsEncryptedObjects = sourceSelectionCriteria.getSseKmsEncryptedObjects();
					rrs.append("\nSSE KMS Encrypted Objects Status=");
					rrs.append(sseKmsEncryptedObjects.getStatus());
					rrs.append("\nReplicationRule Status=");
					rrs.append(replicationRule.getStatus());
				}
				this.xssfHelper.setCell(row, rrs.toString());
			} catch(AmazonS3Exception skip) {
				this.xssfHelper.setCell(row, "");
				this.xssfHelper.setCell(row, "");
			} catch(Exception e) {
				this.xssfHelper.setCell(row, "");
				this.xssfHelper.setCell(row, "");
				e.printStackTrace();
			}
			
			try {			
				StringBuffer bncs = new StringBuffer();
				BucketNotificationConfiguration bucketNotificationConfiguration = this.amazonClients.AmazonS3.getBucketNotificationConfiguration(bucket.getName());
				Map<String, NotificationConfiguration> notiMap = bucketNotificationConfiguration.getConfigurations();
				Set<String> notiKeys = notiMap.keySet();
				Iterator<String> inotiKeys = notiKeys.iterator();
				while(inotiKeys.hasNext()) {
					String key = inotiKeys.next();
					NotificationConfiguration notificationConfiguration = notiMap.get(key);
					Set<String> events = notificationConfiguration.getEvents();
					Iterator<String> ievents = events.iterator();
					
					if(bncs.length() > 0) bncs.append("\n\n");
					
					bncs.append("Events=");
					while(ievents.hasNext()) {
						bncs.append(ievents.next());
						bncs.append("|");
					}
					
					bncs.append("FilterRule");
					
					com.amazonaws.services.s3.model.Filter filter = notificationConfiguration.getFilter();
					S3KeyFilter s3KeyFilter = filter.getS3KeyFilter();
					List<FilterRule> filterRules = s3KeyFilter.getFilterRules();
					for(FilterRule filterRule : filterRules) {
						bncs.append("\nName=");
						bncs.append(filterRule.getName());
						bncs.append(", Value=");
						bncs.append(filterRule.getValue());
					}
				}
				this.xssfHelper.setCell(row, bncs.toString());
			} catch(AmazonS3Exception skip) {
				this.xssfHelper.setCell(row, "");
			} catch(Exception e) {
				this.xssfHelper.setCell(row, "");
				e.printStackTrace();
			}
			
			try {
				BucketWebsiteConfiguration bucketWebsiteConfiguration = this.amazonClients.AmazonS3.getBucketWebsiteConfiguration(bucket.getName());
				this.xssfHelper.setCell(row, bucketWebsiteConfiguration == null ? "" : bucketWebsiteConfiguration.getErrorDocument());
				this.xssfHelper.setCell(row, bucketWebsiteConfiguration == null ? "" : bucketWebsiteConfiguration.getIndexDocumentSuffix());
				
				if(bucketWebsiteConfiguration != null) {
					RedirectRule redirectRule = bucketWebsiteConfiguration.getRedirectAllRequestsTo();
					this.xssfHelper.setCell(row, redirectRule.getHostName());
					this.xssfHelper.setCell(row, redirectRule.getHttpRedirectCode());
					this.xssfHelper.setCell(row, redirectRule.getprotocol());
					this.xssfHelper.setCell(row, redirectRule.getReplaceKeyPrefixWith());
					this.xssfHelper.setCell(row, redirectRule.getReplaceKeyWith());
					
					StringBuffer rors = new StringBuffer();
					List<RoutingRule> routingRules = bucketWebsiteConfiguration.getRoutingRules();
					for(RoutingRule routingRule : routingRules) {
						
						if(rors.length() > 0) rors.append("\n\n");
						
						RoutingRuleCondition routingRuleCondition = routingRule.getCondition();
						
						rors.append("HttpErrorCodeReturnedEquals=");
						rors.append(routingRuleCondition.getHttpErrorCodeReturnedEquals());
						rors.append("\nKeyPrefixEquals=");
						rors.append(routingRuleCondition.getKeyPrefixEquals());
						
						RedirectRule rredirectRule = routingRule.getRedirect();
						rors.append("\nRedirectRule");
						rors.append("\nHostName=");
						rors.append(rredirectRule.getHostName());
						rors.append("\nHttpRedirectCode=");
						rors.append(rredirectRule.getHttpRedirectCode());
						rors.append("\nProtocol=");
						rors.append(rredirectRule.getprotocol());
						rors.append("\nReplaceKeyPrefixWidth=");
						rors.append(rredirectRule.getReplaceKeyPrefixWith());
						rors.append("\nReplaceKeyWidth=");
						rors.append(rredirectRule.getReplaceKeyWith());
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

			} catch(AmazonS3Exception skip) {
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
				AccessControlList accessControlList = this.amazonClients.AmazonS3.getBucketAcl(bucket.getName());
				this.xssfHelper.setCell(row, accessControlList.getOwner().getDisplayName());
				
				StringBuffer grts = new StringBuffer();
				List<Grant> grants = accessControlList.getGrantsAsList();
				for(Grant grant : grants) {
					Grantee grantee = grant.getGrantee();
					
					if(grts.length() > 0) grts.append("\n\n");
					grts.append("Identifier=");
					grts.append(grantee.getIdentifier());
					grts.append("\nTypeIdentifier=");
					grts.append(grantee.getTypeIdentifier());
					grts.append("\nPermission=");
					grts.append(grant.getPermission().getHeaderName());
				}
				this.xssfHelper.setCell(row, grts.toString());
			} catch(AmazonS3Exception skip) {
				this.xssfHelper.setCell(row, "");
				this.xssfHelper.setCell(row, "");
			} catch(Exception e) {
				this.xssfHelper.setCell(row, "");
				this.xssfHelper.setCell(row, "");
				e.printStackTrace();
			}
			
			try {
				StringBuffer lcs = new StringBuffer();
				BucketLifecycleConfiguration bucketLifecycleConfiguration = this.amazonClients.AmazonS3.getBucketLifecycleConfiguration(bucket.getName());
				
				if(bucketLifecycleConfiguration != null) {
					List<com.amazonaws.services.s3.model.BucketLifecycleConfiguration.Rule> rules = bucketLifecycleConfiguration.getRules();
					for(com.amazonaws.services.s3.model.BucketLifecycleConfiguration.Rule rule : rules) {
						AbortIncompleteMultipartUpload abortIncompleteMultipartUpload = rule.getAbortIncompleteMultipartUpload();
						if(lcs.length() > 0) lcs.append("\n\n");
						
						lcs.append("AbortInCompleteMultipartUpload=");
						lcs.append(Integer.toString(abortIncompleteMultipartUpload.getDaysAfterInitiation()));
						lcs.append("\nExpirationDate=");
						lcs.append(rule.getExpirationDate().toString());
						lcs.append("\nExpirationInDays=");
						lcs.append(Integer.toString(rule.getExpirationInDays()));
						lcs.append("\nID=");
						lcs.append(rule.getId());
						lcs.append("\nNonCurrentVersionExpirationInDays=");
						lcs.append(Integer.toString(rule.getNoncurrentVersionExpirationInDays()));
						
						lcs.append("\nNoncurrentVersionTransition");
						List<NoncurrentVersionTransition> noncurrentVersionTransitions = rule.getNoncurrentVersionTransitions();
						for(NoncurrentVersionTransition noncurrentVersionTransition : noncurrentVersionTransitions) {
							lcs.append("\nDays=");
							lcs.append(Integer.toString(noncurrentVersionTransition.getDays()));
							lcs.append("\nStorageClass=");
							lcs.append(noncurrentVersionTransition.getStorageClassAsString());
						}
						
						lcs.append("\nStatus=");
						lcs.append(rule.getStatus());
						
						lcs.append("\nTransitions");
						List<Transition> transitions = rule.getTransitions();
						for(Transition transition : transitions) {
							lcs.append("\nDate=");
							lcs.append(transition.getDate().toString());
							lcs.append("\nDays=");
							lcs.append(Integer.toString(transition.getDays()));
							lcs.append("\nStorageClass=");
							lcs.append(transition.getStorageClassAsString());
						}
					}
				}
				this.xssfHelper.setRightThinCell(row, lcs.toString());
			} catch(AmazonS3Exception skip) {
				this.xssfHelper.setRightThinCell(row, "");
			} catch(Exception e) {
				this.xssfHelper.setRightThinCell(row, "");
				e.printStackTrace();
			}
		}
	}
	
	
	private void makeACM() {
		
		XSSFRow row = null;
		
		ListCertificatesResult listCertificatesResult = this.amazonClients.AwsCertificateManager.listCertificates(new ListCertificatesRequest());
		List<CertificateSummary> certificateSummaries = listCertificatesResult.getCertificateSummaryList();
		for (CertificateSummary certificateSummary : certificateSummaries) {
			
			DescribeCertificateResult describeCertificateResult = this.amazonClients.AwsCertificateManager.describeCertificate(new DescribeCertificateRequest().withCertificateArn(certificateSummary.getCertificateArn()));
			CertificateDetail certificateDetail = describeCertificateResult.getCertificate();
			
			row = this.xssfHelper.createRow(this.acmSheet, 1);			
			this.xssfHelper.setLeftThinCell(row, Integer.toString(row.getRowNum() - 1));
			this.xssfHelper.setCell(row, certificateDetail.getCertificateArn());
			this.xssfHelper.setCell(row, certificateDetail.getCertificateAuthorityArn());
			this.xssfHelper.setCell(row, certificateDetail.getDomainName());
			this.xssfHelper.setCell(row, certificateDetail.getStatus());
			this.xssfHelper.setCell(row, certificateDetail.getSubject());
			
			StringBuffer sans = new StringBuffer();
			List<String> subjectAlternativeNames = certificateDetail.getSubjectAlternativeNames();
			for(String subjectAlternativeName : subjectAlternativeNames) {
				if(sans.length() > 0) sans.append("\n");
				sans.append(subjectAlternativeName);
			}
			this.xssfHelper.setCell(row, sans.toString());
			this.xssfHelper.setCell(row, certificateDetail.getType());
			this.xssfHelper.setCell(row, certificateDetail.getKeyAlgorithm());
			
			StringBuffer kus = new StringBuffer();
			List<KeyUsage> keyUsages = certificateDetail.getKeyUsages();
			for(KeyUsage keyUsage : keyUsages) {
				if(kus.length() > 0) kus.append("\n");
				kus.append(keyUsage.getName());
			}
			
			this.xssfHelper.setCell(row, kus.toString());
			this.xssfHelper.setCell(row, certificateDetail.getSerial());
			this.xssfHelper.setCell(row, certificateDetail.getSignatureAlgorithm());
			this.xssfHelper.setCell(row, certificateDetail.getOptions().getCertificateTransparencyLoggingPreference());
			
			this.xssfHelper.setCell(row, this.getDomainValidation(certificateDetail.getDomainValidationOptions()));
			
			StringBuffer ekus = new StringBuffer();
			List<ExtendedKeyUsage> extendedKeyUsages = certificateDetail.getExtendedKeyUsages();
			for(ExtendedKeyUsage extendedKeyUsage : extendedKeyUsages) {
				if(ekus.length() > 0) ekus.append("\n");
				ekus.append(extendedKeyUsage.getName());
				ekus.append(" (");
				ekus.append(extendedKeyUsage.getOID());
				ekus.append(")");
			}
			this.xssfHelper.setCell(row, ekus.toString());
			this.xssfHelper.setCell(row, certificateDetail.getFailureReason());
			this.xssfHelper.setCell(row, certificateDetail.getIssuer());
			this.xssfHelper.setCell(row, certificateDetail.getRenewalEligibility());
			RenewalSummary renewalSummary = certificateDetail.getRenewalSummary();
			this.xssfHelper.setCell(row, renewalSummary == null ? "" : renewalSummary.getRenewalStatus());
			this.xssfHelper.setCell(row, renewalSummary == null ? "" : this.getDomainValidation(renewalSummary.getDomainValidationOptions()));
			this.xssfHelper.setCell(row, certificateDetail.getCreatedAt() == null ? "" : certificateDetail.getCreatedAt().toString());
			this.xssfHelper.setCell(row, certificateDetail.getImportedAt() == null ? "" : certificateDetail.getImportedAt().toString());
			this.xssfHelper.setCell(row, certificateDetail.getInUseBy() == null ? "" : certificateDetail.getInUseBy().toString());
			this.xssfHelper.setCell(row, certificateDetail.getIssuedAt() == null ? "" : certificateDetail.getIssuedAt().toString());
			this.xssfHelper.setCell(row, certificateDetail.getRevokedAt() == null ? "" : certificateDetail.getRevokedAt().toString());
			this.xssfHelper.setCell(row, certificateDetail.getRevocationReason());
			this.xssfHelper.setCell(row, certificateDetail.getNotAfter() == null ? "" : certificateDetail.getNotAfter().toString());
			this.xssfHelper.setCell(row, certificateDetail.getNotBefore() == null ? "" : certificateDetail.getNotBefore().toString());
			
		}
	}
	
	private String getDomainValidation(List<DomainValidation> domainValidations) {
	
		StringBuffer dvs = new StringBuffer();
		for(DomainValidation domainValidation : domainValidations) {
			if(dvs.length() > 0) dvs.append("\n\n");
			dvs.append("DomainName=");
			dvs.append(domainValidation.getDomainName());
			dvs.append("\nResourceRecord=");
			dvs.append(domainValidation.getResourceRecord());
			dvs.append("\nValidation Domain=");
			dvs.append(domainValidation.getValidationDomain());
			dvs.append("\nValidation Method=");
			dvs.append(domainValidation.getValidationMethod());

			List<String> validationEmails = domainValidation.getValidationEmails();
			if(validationEmails != null) {
				dvs.append("\nValidationEmail");
				for(String validationEmail : validationEmails) {
					dvs.append("\n\t-");
					dvs.append(validationEmail);
				}
			}
			dvs.append("\nValidationStatus=");
			dvs.append(domainValidation.getValidationStatus());
		}
		
		return dvs.toString();
	}
	
	private void makeKMS() {
		
		XSSFRow row = null;
		
		Map<String, AliasListEntry> aliasMap = new HashMap<>();
		ListAliasesResult listAliasesResult = this.amazonClients.AwsKMS.listAliases();
		List<AliasListEntry> aliasListEntries = listAliasesResult.getAliases();
		for (AliasListEntry aliasListEntry : aliasListEntries) {
			
			if(aliasListEntry.getTargetKeyId() == null) {
				row = this.xssfHelper.createRow(this.kmsSheet, 1);			
				this.xssfHelper.setLeftThinCell(row, Integer.toString(row.getRowNum() - 1));
				this.xssfHelper.setCell(row, "");
				this.xssfHelper.setCell(row, aliasListEntry.getAliasArn());
				this.xssfHelper.setCell(row, aliasListEntry.getAliasName());
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
				aliasMap.put(aliasListEntry.getTargetKeyId(), aliasListEntry);
			}
			
		}
		
		ListKeysResult listKeysResult = this.amazonClients.AwsKMS.listKeys();
		List<KeyListEntry> keyListEntries = listKeysResult.getKeys();
		for (KeyListEntry keyListEntry : keyListEntries) {
			
			row = this.xssfHelper.createRow(this.kmsSheet, 1);
			
			AliasListEntry aliasListEntry = aliasMap.get(keyListEntry.getKeyId());
			
			DescribeKeyResult describeKeyResult = this.amazonClients.AwsKMS.describeKey(new DescribeKeyRequest().withKeyId(keyListEntry.getKeyId()));
			KeyMetadata keyMetadata = describeKeyResult.getKeyMetadata();
			
			this.xssfHelper.setLeftThinCell(row, Integer.toString(row.getRowNum() - 1));
			this.xssfHelper.setCell(row, keyMetadata.getKeyManager());
			this.xssfHelper.setCell(row, aliasListEntry == null ? "" : aliasListEntry.getAliasArn());
			this.xssfHelper.setCell(row, aliasListEntry == null ? "" : aliasListEntry.getAliasName());

			this.xssfHelper.setCell(row, keyMetadata.getKeyId());
			this.xssfHelper.setCell(row, keyMetadata.getArn());
			this.xssfHelper.setCell(row, keyMetadata.getAWSAccountId());
			this.xssfHelper.setCell(row, keyMetadata.getCreationDate().toString());
			this.xssfHelper.setCell(row, keyMetadata.getDeletionDate() == null ? "" : keyMetadata.getDeletionDate().toString());
			this.xssfHelper.setCell(row, keyMetadata.getDescription());
			this.xssfHelper.setCell(row, keyMetadata.getEnabled().toString());
			this.xssfHelper.setCell(row, keyMetadata.getExpirationModel());
			this.xssfHelper.setCell(row, keyMetadata.getKeyState());
			this.xssfHelper.setCell(row, keyMetadata.getKeyUsage());
			this.xssfHelper.setCell(row, keyMetadata.getOrigin());
			this.xssfHelper.setRightThinCell(row, keyMetadata.getValidTo() == null ? "" : keyMetadata.getValidTo().toString());
				
		}
		
	}
	
	private void makeElasticCache() {
		XSSFRow row = null;
    
		List<ReplicationGroup> replicationGroups = this.amazonClients.AmazonElastiCache.describeReplicationGroups().getReplicationGroups();
		for (ReplicationGroup replicationGroup : replicationGroups) {
			int firstRowNum = 0;
			row = null;
			List<NodeGroup> nodeGroups = replicationGroup.getNodeGroups();
			for (NodeGroup nodeGroup : nodeGroups) {
        
				List<NodeGroupMember> nodeGroupMembers = nodeGroup.getNodeGroupMembers();
				for (NodeGroupMember nodeGroupMember : nodeGroupMembers) {
					CacheCluster cacheCluster = (CacheCluster)this.amazonClients.AmazonElastiCache.describeCacheClusters(new DescribeCacheClustersRequest().withCacheClusterId(nodeGroupMember.getCacheClusterId())).getCacheClusters().get(0);
          
					row = this.xssfHelper.createRow(this.elastiCacheSheet, 1);
					if (firstRowNum == 0) {
						firstRowNum = row.getRowNum();
					}
					this.xssfHelper.setLeftThinCell(row, Integer.toString(row.getRowNum() - 2));
					this.xssfHelper.setCell(row, this.region.getName());
					this.xssfHelper.setCell(row, replicationGroup.getReplicationGroupId());
					this.xssfHelper.setCell(row, cacheCluster.getEngine());
					this.xssfHelper.setCell(row, cacheCluster.getEngineVersion());
					this.xssfHelper.setCell(row, nodeGroup.getPrimaryEndpoint().getAddress() + ":" + nodeGroup.getPrimaryEndpoint().getPort());
					this.xssfHelper.setCell(row, cacheCluster.getCacheParameterGroup().getCacheParameterGroupName() + "(" + cacheCluster.getCacheParameterGroup().getParameterApplyStatus() + ")");
					this.xssfHelper.setCell(row, cacheCluster.getCacheSubnetGroupName());
          
					StringBuffer securityGroups = new StringBuffer();
					List<SecurityGroupMembership> securityGroupMemberships = cacheCluster.getSecurityGroups();
					for (SecurityGroupMembership securityGroupMembership : securityGroupMemberships) {
						if (securityGroups.length() > 0) {
							securityGroups.append("\n");
						}
						securityGroups.append(securityGroupMembership.getSecurityGroupId());
						securityGroups.append(" (");
						securityGroups.append(securityGroupMembership.getStatus());
						securityGroups.append(")");
					}
					this.xssfHelper.setCell(row, securityGroups.toString());
					this.xssfHelper.setCell(row, cacheCluster.getPreferredMaintenanceWindow());
          
					this.xssfHelper.setCell(row, cacheCluster.getCacheClusterId());
					this.xssfHelper.setCell(row, nodeGroupMember.getPreferredAvailabilityZone());
					this.xssfHelper.setCell(row, nodeGroupMember.getCurrentRole());
					this.xssfHelper.setCell(row, cacheCluster.getCacheNodeType());
					this.xssfHelper.setRightThinCell(row, nodeGroupMember.getReadEndpoint().getAddress() + ":" + nodeGroupMember.getReadEndpoint().getPort());
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
		
		DescribeDBClustersResult describeDBClustersResult = this.amazonClients.AmazonRDS.describeDBClusters();
		List<DBCluster> dbClusters = describeDBClustersResult.getDBClusters();
		for (DBCluster dbCluster : dbClusters) {

			row = this.xssfHelper.createRow(this.rdsClusterSheet, 1);
			this.xssfHelper.setLeftThinCell(row, Integer.toString(row.getRowNum() - 1));
			this.xssfHelper.setCell(row, "Cluster");
			this.xssfHelper.setCell(row, dbCluster.getDBClusterIdentifier());
			
			StringBuffer avzs = new StringBuffer();
			List<String> availabilityZones = dbCluster.getAvailabilityZones();
			for (String availabilityZone : availabilityZones) {
				if(avzs.length() > 0) avzs.append("\n");
				avzs.append(availabilityZone);
			}
			this.xssfHelper.setCell(row, avzs.toString());
			this.xssfHelper.setCell(row, dbCluster.getDBSubnetGroup());
			
			StringBuffer dbclopgs = new StringBuffer();
			List<DBClusterOptionGroupStatus> dbClusterOptionGroupStatList = dbCluster.getDBClusterOptionGroupMemberships();
			for(DBClusterOptionGroupStatus dbClusterOptionGroupStatus : dbClusterOptionGroupStatList) {
				if(dbclopgs.length() > 0) dbclopgs.append("\n");
				dbclopgs.append(dbClusterOptionGroupStatus.getDBClusterOptionGroupName());
				dbclopgs.append(" (");
				dbclopgs.append(dbClusterOptionGroupStatus.getStatus());
				dbclopgs.append(")");
			}
			this.xssfHelper.setCell(row, dbclopgs.toString());
			this.xssfHelper.setCell(row, dbCluster.getDBClusterParameterGroup());
			this.xssfHelper.setCell(row, dbCluster.getDbClusterResourceId());
			
			this.xssfHelper.setCell(row, dbCluster.getDatabaseName());
			this.xssfHelper.setCell(row, dbCluster.getCharacterSetName());

			this.xssfHelper.setCell(row, dbCluster.getAllocatedStorage().toString());
			this.xssfHelper.setCell(row, dbCluster.getCloneGroupId());
			
			this.xssfHelper.setCell(row, dbCluster.getBacktrackWindow() == null ? "" : dbCluster.getBacktrackWindow().toString());
			this.xssfHelper.setCell(row, dbCluster.getBackupRetentionPeriod() == null ? "" : dbCluster.getBackupRetentionPeriod().toString());
			this.xssfHelper.setCell(row, dbCluster.getBacktrackConsumedChangeRecords() == null ? "" : dbCluster.getBacktrackConsumedChangeRecords().toString());

			this.xssfHelper.setCell(row, dbCluster.getDBClusterArn());
			
			StringBuffer dbclrs = new StringBuffer();
			List<DBClusterRole> dbClusterRoles = dbCluster.getAssociatedRoles();
			for(DBClusterRole dbClusterRole : dbClusterRoles) {
				if(dbclrs.length() > 0) dbclrs.append("\n");
				dbclrs.append("(");
				dbclrs.append(dbClusterRole.getStatus());
				dbclrs.append(") ");
				dbclrs.append(dbClusterRole.getRoleArn());
			}
			this.xssfHelper.setCell(row, dbclrs.toString());
			
			this.xssfHelper.setCell(row, dbCluster.getEarliestBacktrackTime() == null ? "" : dbCluster.getEarliestBacktrackTime().toString());
			this.xssfHelper.setCell(row, dbCluster.getEarliestRestorableTime() == null ? "" : dbCluster.getEarliestRestorableTime().toString());
			
			StringBuffer ecle = new StringBuffer();
			List<String> enabledCloudwatchLogsExports = dbCluster.getEnabledCloudwatchLogsExports();
			for (String enabledCloudwatchLogsExport : enabledCloudwatchLogsExports) {
				if(ecle.length() > 0) ecle.append("\n");
				ecle.append(enabledCloudwatchLogsExport);
			}
			this.xssfHelper.setCell(row, ecle.toString());
			this.xssfHelper.setCell(row, dbCluster.getEndpoint());
			this.xssfHelper.setCell(row, dbCluster.getEngine());
			this.xssfHelper.setCell(row, dbCluster.getEngineVersion());
			
			this.xssfHelper.setCell(row, dbCluster.getHostedZoneId());
			this.xssfHelper.setCell(row, dbCluster.getIAMDatabaseAuthenticationEnabled().toString());
			this.xssfHelper.setCell(row, dbCluster.getKmsKeyId());
			this.xssfHelper.setCell(row, dbCluster.getLatestRestorableTime().toString());
			this.xssfHelper.setCell(row, dbCluster.getMasterUsername());
			this.xssfHelper.setCell(row, dbCluster.getMultiAZ().toString());
			this.xssfHelper.setCell(row, dbCluster.getPercentProgress());
			this.xssfHelper.setCell(row, dbCluster.getPort().toString());
			this.xssfHelper.setCell(row, dbCluster.getPreferredBackupWindow());
			this.xssfHelper.setCell(row, dbCluster.getPreferredMaintenanceWindow());
			this.xssfHelper.setCell(row, dbCluster.getReaderEndpoint());
			
			StringBuffer rrif = new StringBuffer();
			List<String> readReplicaIdentifiers = dbCluster.getReadReplicaIdentifiers();
			for(String readReplicaIdentifier : readReplicaIdentifiers) {
				if(rrif.length() > 0) rrif.append("\n");
				rrif.append(readReplicaIdentifier);
			}
			this.xssfHelper.setCell(row, rrif.toString());
			this.xssfHelper.setCell(row, dbCluster.getReplicationSourceIdentifier());
			this.xssfHelper.setCell(row, dbCluster.getStatus());
			this.xssfHelper.setCell(row, dbCluster.getStorageEncrypted().toString());
			
			StringBuffer vsgs = new StringBuffer();
			List<VpcSecurityGroupMembership> vpcSecurityGroupMemberships = dbCluster.getVpcSecurityGroups();
			for(VpcSecurityGroupMembership vpcSecurityGroupMembership : vpcSecurityGroupMemberships) {
				if(vsgs.length() > 0) vsgs.append("\n");
				vsgs.append(vpcSecurityGroupMembership.getVpcSecurityGroupId());
				vsgs.append(" (");
				vsgs.append(vpcSecurityGroupMembership.getStatus());
				vsgs.append(")");
			}
			this.xssfHelper.setCell(row, vsgs.toString());
			
			StringBuffer dcms = new StringBuffer();
			List<DBClusterMember> dbClusterMembers = dbCluster.getDBClusterMembers();
			for (DBClusterMember dbClusterMember : dbClusterMembers) {
				if(dcms.length() > 0) dcms.append("\n");
				dcms.append("DBInstanceidentifier=");
				dcms.append(dbClusterMember.getDBInstanceIdentifier());
				dcms.append(", IsClusterWriter=");
				dcms.append(dbClusterMember.getIsClusterWriter().toString());
				dcms.append(", DBClusterParameterGroupStatus=");
				dcms.append(dbClusterMember.getDBClusterParameterGroupStatus());
				dcms.append(", PromotionTier");
				dcms.append(dbClusterMember.getPromotionTier().toString());
			}
			this.xssfHelper.setLeftRightThinCell(row, dcms.toString());

		}
		
	}
	
	
	private void makeRDSInstance() {
		
		DescribeDBInstancesResult describeDBInstancesResult = this.amazonClients.AmazonRDS.describeDBInstances();
		List<DBInstance> dbInstances = describeDBInstancesResult.getDBInstances();
		for (DBInstance dbInstance : dbInstances) {
			this.makeRDSDBInstances(dbInstance);
		}
		
	}
	
	private void makeRDSDBInstances(DBInstance dbInstance) {
		
		XSSFRow row = null;
		
		row = this.xssfHelper.createRow(this.rdsInstanceSheet, 1);
		this.xssfHelper.setLeftThinCell(row, Integer.toString(row.getRowNum() - 1));
		this.xssfHelper.setCell(row, dbInstance.getDBClusterIdentifier() == null ? "Instance" : "Cluster Member");
		this.xssfHelper.setCell(row, dbInstance.getDBClusterIdentifier());
		this.xssfHelper.setCell(row, dbInstance.getDBInstanceIdentifier());
		this.xssfHelper.setCell(row, dbInstance.getDBInstanceClass());
		this.xssfHelper.setCell(row, dbInstance.getDbInstancePort().toString());
		this.xssfHelper.setCell(row, dbInstance.getDBInstanceStatus());
		this.xssfHelper.setCell(row, dbInstance.getAvailabilityZone());
		this.xssfHelper.setCell(row, dbInstance.getDBSubnetGroup().getDBSubnetGroupName());
		
		StringBuffer dbsgms = new StringBuffer();
		List<DBSecurityGroupMembership> dbSecurityGroups = dbInstance.getDBSecurityGroups();
		for (DBSecurityGroupMembership dbSecurityGroupMembership : dbSecurityGroups) {
			if(dbsgms.length() > 0) dbsgms.append("\n");
			dbsgms.append(dbSecurityGroupMembership.getDBSecurityGroupName());
			dbsgms.append(" (");
			dbsgms.append(dbSecurityGroupMembership.getStatus());
			dbsgms.append(")");
		}
		
		this.xssfHelper.setCell(row, dbsgms.toString());
		
		StringBuffer dbpgs = new StringBuffer();
		List<DBParameterGroupStatus> dbParameterGroupStatList = dbInstance.getDBParameterGroups();
		for (DBParameterGroupStatus dbParameterGroupStatus : dbParameterGroupStatList) {
			if(dbpgs.length() > 0) dbpgs.append("\n");
			dbpgs.append(dbParameterGroupStatus.getDBParameterGroupName());
			dbpgs.append(" (");
			dbpgs.append(dbParameterGroupStatus.getParameterApplyStatus());
			dbpgs.append(")");
		}
		
		this.xssfHelper.setCell(row, dbpgs.toString());
		this.xssfHelper.setCell(row, dbInstance.getDBName());
		this.xssfHelper.setCell(row, dbInstance.getCharacterSetName());

		this.xssfHelper.setCell(row, dbInstance.getCACertificateIdentifier());
		
		this.xssfHelper.setCell(row, dbInstance.getAutoMinorVersionUpgrade().toString());
		this.xssfHelper.setCell(row, dbInstance.getBackupRetentionPeriod().toString());

		this.xssfHelper.setCell(row, dbInstance.getCopyTagsToSnapshot().toString());

		this.xssfHelper.setCell(row, dbInstance.getDbiResourceId());
		this.xssfHelper.setCell(row, dbInstance.getAllocatedStorage().toString());
		this.xssfHelper.setCell(row, dbInstance.getDBInstanceArn());
		
		StringBuffer dmss = new StringBuffer();
		List<DomainMembership> domainMemberships = dbInstance.getDomainMemberships();
		for (DomainMembership domainMembership : domainMemberships) {
			if(dmss.length() > 0) dmss.append("\n");
			dmss.append("Domain=");
			dmss.append(domainMembership.getDomain());
			dmss.append(", FQDN=");
			dmss.append(domainMembership.getFQDN());
			dmss.append(", IAMRole=");
			dmss.append(domainMembership.getIAMRoleName());
			dmss.append(", Status=");
			dmss.append(domainMembership.getStatus());
		}
		this.xssfHelper.setCell(row, dmss.toString());
		
		StringBuffer ecle = new StringBuffer();
		List<String> enabledCloudwatchLogsExports = dbInstance.getEnabledCloudwatchLogsExports();
		for (String enabledCloudwatchLogsExport : enabledCloudwatchLogsExports) {
			if(ecle.length() > 0) ecle.append("\n");
			ecle.append(enabledCloudwatchLogsExport);
		}
		this.xssfHelper.setCell(row, ecle.toString());
		
		Endpoint endpoint = dbInstance.getEndpoint();
		this.xssfHelper.setCell(row, endpoint.getAddress());
		this.xssfHelper.setCell(row, endpoint.getHostedZoneId());
		this.xssfHelper.setCell(row, endpoint.getPort().toString());
		this.xssfHelper.setCell(row, dbInstance.getEngine());
		this.xssfHelper.setCell(row, dbInstance.getEngineVersion());
		this.xssfHelper.setCell(row, dbInstance.getEnhancedMonitoringResourceArn());
		this.xssfHelper.setCell(row, dbInstance.getIAMDatabaseAuthenticationEnabled().toString());
		this.xssfHelper.setCell(row, dbInstance.getInstanceCreateTime().toString());
		this.xssfHelper.setCell(row, dbInstance.getIops() == null ? "" : dbInstance.getIops().toString());
		this.xssfHelper.setCell(row, dbInstance.getKmsKeyId());
		this.xssfHelper.setCell(row, dbInstance.getLatestRestorableTime() == null ? "" : dbInstance.getLatestRestorableTime().toString());
		this.xssfHelper.setCell(row, dbInstance.getLicenseModel());
		this.xssfHelper.setCell(row, dbInstance.getMasterUsername());
		this.xssfHelper.setCell(row, dbInstance.getMonitoringInterval() == null ? "" : dbInstance.getMonitoringInterval().toString());
		this.xssfHelper.setCell(row, dbInstance.getMonitoringRoleArn());
		this.xssfHelper.setCell(row, dbInstance.getMultiAZ().toString());
		this.xssfHelper.setCell(row, dbInstance.getOptionGroupMemberships().toString());
		this.xssfHelper.setCell(row, dbInstance.getPerformanceInsightsEnabled().toString());
		this.xssfHelper.setCell(row, dbInstance.getPerformanceInsightsKMSKeyId());
		this.xssfHelper.setCell(row, dbInstance.getPerformanceInsightsRetentionPeriod() == null ? "" : dbInstance.getPerformanceInsightsRetentionPeriod().toString());
		this.xssfHelper.setCell(row, dbInstance.getPreferredBackupWindow());
		this.xssfHelper.setCell(row, dbInstance.getPreferredMaintenanceWindow());
		this.xssfHelper.setCell(row, dbInstance.getProcessorFeatures().toString());
		this.xssfHelper.setCell(row, dbInstance.getPromotionTier() == null ? "" : dbInstance.getPromotionTier().toString());
		this.xssfHelper.setCell(row, dbInstance.getPubliclyAccessible().toString());
		
		StringBuffer rdcis = new StringBuffer();
		List<String> readReplicaDBClusterIdentifiers = dbInstance.getReadReplicaDBClusterIdentifiers();
		for(String readReplicaDBClusterIdentifier : readReplicaDBClusterIdentifiers) {
			if(rdcis.length() > 0) rdcis.append("\n");
			rdcis.append(readReplicaDBClusterIdentifier);
		}
		this.xssfHelper.setCell(row, rdcis.toString());
		
		StringBuffer rdiis = new StringBuffer();
		List<String> readReplicaDBInstanceIdentifiers = dbInstance.getReadReplicaDBInstanceIdentifiers();
		for(String readReplicaDBInstanceIdentifier : readReplicaDBInstanceIdentifiers) {
			if(rdiis.length() > 0) rdiis.append("\n");
			rdiis.append(readReplicaDBInstanceIdentifier);
		}
		
		this.xssfHelper.setCell(row, rdiis.toString());
		this.xssfHelper.setCell(row, dbInstance.getReadReplicaSourceDBInstanceIdentifier());
		this.xssfHelper.setCell(row, dbInstance.getSecondaryAvailabilityZone());
		this.xssfHelper.setCell(row, dbInstance.getStatusInfos().toString());
		this.xssfHelper.setCell(row, dbInstance.getStorageEncrypted().toString());
		this.xssfHelper.setCell(row, dbInstance.getStorageType());
		this.xssfHelper.setCell(row, dbInstance.getTdeCredentialArn());
		this.xssfHelper.setCell(row, dbInstance.getTimezone());
		
		StringBuffer vsgs = new StringBuffer();
		List<VpcSecurityGroupMembership> vpcSecurityGroupMemberships = dbInstance.getVpcSecurityGroups();
		for(VpcSecurityGroupMembership vpcSecurityGroupMembership : vpcSecurityGroupMemberships) {
			if(vsgs.length() > 0) vsgs.append("\n");
			vsgs.append(vpcSecurityGroupMembership.getVpcSecurityGroupId());
			vsgs.append(" (");
			vsgs.append(vpcSecurityGroupMembership.getStatus());
			vsgs.append(")");
		}
		this.xssfHelper.setRightThinCell(row, vsgs.toString());
	}
	
	
	private void makeEBS() {
		XSSFRow row = null;
    
		List<Volume> volumes = this.amazonClients.AmazonEC2.describeVolumes().getVolumes();
		for (Volume volume : volumes) {
			row = this.xssfHelper.createRow(this.ebsSheet, 1);
			this.xssfHelper.setLeftThinCell(row, Integer.toString(row.getRowNum() - 1));
			this.xssfHelper.setCell(row, this.region.getName());
			this.xssfHelper.setCell(row, getNameTagValue(volume.getTags()));
			this.xssfHelper.setCell(row, volume.getVolumeId());
			this.xssfHelper.setCell(row, volume.getVolumeType());
			this.xssfHelper.setCell(row, volume.getSize().toString());
			this.xssfHelper.setCell(row, volume.getIops() == null ? "-" : volume.getIops().toString());
			this.xssfHelper.setCell(row, volume.getAvailabilityZone());
			this.xssfHelper.setCell(row, volume.getSnapshotId());
      
			StringBuffer volumeAttachmentBuffer = new StringBuffer();
			List<VolumeAttachment> volumeAttachments = volume.getAttachments();
			for (VolumeAttachment volumeAttachment : volumeAttachments) {
				com.amazonaws.services.ec2.model.Instance instance = (com.amazonaws.services.ec2.model.Instance)((Reservation)this.amazonClients.AmazonEC2.describeInstances(new DescribeInstancesRequest().withInstanceIds(new String[] { volumeAttachment.getInstanceId() })).getReservations().get(0)).getInstances().get(0);
				if (volumeAttachmentBuffer.length() > 0) {
					volumeAttachmentBuffer.append("\n");
				}
				volumeAttachmentBuffer.append(volumeAttachment.getInstanceId());
				volumeAttachmentBuffer.append("(");
				volumeAttachmentBuffer.append(getNameTagValue(instance.getTags()));
				volumeAttachmentBuffer.append("):");
				volumeAttachmentBuffer.append(volumeAttachment.getDevice());
				volumeAttachmentBuffer.append("(");
				volumeAttachmentBuffer.append(volumeAttachment.getState());
				volumeAttachmentBuffer.append(")");
			}
			this.xssfHelper.setCell(row, volumeAttachmentBuffer.toString());
      
			this.xssfHelper.setCell(row, volume.getEncrypted().booleanValue() ? "Ecrypted" : "Not Encrypted");
			this.xssfHelper.setCell(row, volume.getKmsKeyId());
			this.xssfHelper.setRightThinCell(row, this.getAllTagValue(volume.getTags()));
		}
	}
  
	
	private void makeClassicELB() {
		XSSFRow row = null;
    
		List<LoadBalancerDescription> loadBalancerDescriptions = this.amazonClients.AmazonElasticLoadBalancing.describeLoadBalancers().getLoadBalancerDescriptions();
		for (LoadBalancerDescription loadBalancerDescription : loadBalancerDescriptions) {
			row = this.xssfHelper.createRow(this.classicElbSheet, 1);
			int startRowNum = row.getRowNum();
			this.xssfHelper.setLeftThinCell(row, Integer.toString(row.getRowNum() - 1));
			this.xssfHelper.setCell(row, this.region.getName());
			this.xssfHelper.setCell(row, loadBalancerDescription.getVPCId());
			this.xssfHelper.setCell(row, "");
			StringBuffer avzBuffer = new StringBuffer();
			List<String> subnetIds = loadBalancerDescription.getSubnets();
			for (String subnetId : subnetIds) {
				Subnet subnet = this.amazonClients.AmazonEC2.describeSubnets(new DescribeSubnetsRequest().withSubnetIds(new String[] { subnetId })).getSubnets().get(0);
				if (avzBuffer.length() > 0) {
					avzBuffer.append("\n");
				}
				avzBuffer.append(subnet.getAvailabilityZone());
				avzBuffer.append("(");
				avzBuffer.append(subnet.getSubnetId());
				avzBuffer.append(")");
			}
			this.xssfHelper.setCell(row, avzBuffer.toString());
			this.xssfHelper.setCell(row, loadBalancerDescription.getLoadBalancerName());
			this.xssfHelper.setCell(row, loadBalancerDescription.getDNSName());
      
			StringBuffer instanceBuffer = new StringBuffer();
			List<com.amazonaws.services.elasticloadbalancing.model.Instance> instances = loadBalancerDescription.getInstances();
			for (com.amazonaws.services.elasticloadbalancing.model.Instance instance : instances) {
				Reservation reservation = (Reservation)this.amazonClients.AmazonEC2.describeInstances(new DescribeInstancesRequest().withInstanceIds(new String[] { instance.getInstanceId() })).getReservations().get(0);
				com.amazonaws.services.ec2.model.Instance ec2Instance = (com.amazonaws.services.ec2.model.Instance)reservation.getInstances().get(0);
				if (instanceBuffer.length() > 0) {
					instanceBuffer.append("\n");
				}
				instanceBuffer.append(ec2Instance.getInstanceId() + " [" + getNameTagValue(ec2Instance.getTags()) + "]");
			}

			this.xssfHelper.setCell(row, instanceBuffer.toString());
			this.xssfHelper.setSubHeadLeftThinCell(row, "Load Balancer Protocol");
			this.xssfHelper.setSubHeadCell(row, "Load Balancer Port");
			this.xssfHelper.setSubHeadCell(row, "Instance Protocol");
			this.xssfHelper.setSubHeadRightThinCell(row, "Instance Port");
			this.xssfHelper.setRightThinCell(row, loadBalancerDescription.getHealthCheck().getTarget());
      
			List<ListenerDescription> listenerDescriptions = loadBalancerDescription.getListenerDescriptions();
			for (ListenerDescription listenerDescription : listenerDescriptions) {
				Listener listener = listenerDescription.getListener();
        
				row = this.xssfHelper.createRow(this.classicElbSheet, 1);
				this.xssfHelper.setLeftThinCell(row, Integer.toString(row.getRowNum() - 1));
				this.xssfHelper.setCell(row, "");
				this.xssfHelper.setCell(row, "");
				this.xssfHelper.setCell(row, "");
				this.xssfHelper.setCell(row, "");
				this.xssfHelper.setCell(row, "");
				this.xssfHelper.setCell(row, "");
				this.xssfHelper.setCell(row, "");
				this.xssfHelper.setLeftThinCell(row, listener.getProtocol());
				this.xssfHelper.setCell(row, listener.getLoadBalancerPort().toString());
				this.xssfHelper.setCell(row, listener.getInstanceProtocol());
				this.xssfHelper.setRightThinCell(row, listener.getInstancePort().toString());
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
	  	  
		com.amazonaws.services.elasticloadbalancingv2.model.DescribeLoadBalancersResult describeLoadBalancersResult = this.amazonClients.AmazonElasticLoadBalancing2.describeLoadBalancers(new DescribeLoadBalancersRequest());
		List<LoadBalancer> loadBalancers = describeLoadBalancersResult.getLoadBalancers();
		for (LoadBalancer loadBalancer : loadBalancers) {
		  
			row = this.xssfHelper.createRow(this.otherElbSheet, 1);
			//int startRowNum = row.getRowNum();
			this.xssfHelper.setLeftThinCell(row, Integer.toString(row.getRowNum() - 1));
			this.xssfHelper.setCell(row, this.region.getName());
			this.xssfHelper.setCell(row, loadBalancer.getVpcId());
	      
			StringBuffer azs = new StringBuffer();
			List<AvailabilityZone> availabilityZones = loadBalancer.getAvailabilityZones();
			for (AvailabilityZone availabilityZone : availabilityZones) {
				if(azs.length() > 0) azs.append("\n");
				azs.append("Zone Name = ");
				azs.append(availabilityZone.getZoneName());
				azs.append("\n Subnet ID = ");
				azs.append(availabilityZone.getSubnetId());
			
				StringBuffer lba = new StringBuffer();
				List<LoadBalancerAddress> loadBalancerAddresses = availabilityZone.getLoadBalancerAddresses();
				if(loadBalancerAddresses != null) {
					for(LoadBalancerAddress loadBalancerAddress : loadBalancerAddresses) {
						lba.append("\nIp : ");
						lba.append(loadBalancerAddress.getIpAddress());
					}
				}
				azs.append(lba.toString() + "");
			}
	      
			this.xssfHelper.setCell(row, azs.toString());
			this.xssfHelper.setCell(row, loadBalancer.getType());
			this.xssfHelper.setCell(row, loadBalancer.getScheme());
			this.xssfHelper.setCell(row, loadBalancer.getIpAddressType());
			this.xssfHelper.setCell(row, loadBalancer.getLoadBalancerName());
			this.xssfHelper.setCell(row, loadBalancer.getCanonicalHostedZoneId());
			String stsReason = loadBalancer.getState().getReason();
			this.xssfHelper.setCell(row, loadBalancer.getState().getCode() + (stsReason == null ? "" : "(" + stsReason + ")"));
	      
			StringBuffer sgs = new StringBuffer();
			List<String> securityGroups = loadBalancer.getSecurityGroups();
			if(securityGroups != null) {
				for(String securityGroup : securityGroups) {
					sgs.append("\n");
					sgs.append(securityGroup);
				}
			}

			this.xssfHelper.setCell(row, sgs.toString() + "");
			this.xssfHelper.setCell(row, loadBalancer.getDNSName());
			this.xssfHelper.setCell(row, "");
			this.xssfHelper.setCell(row, "");
			this.xssfHelper.setCell(row, loadBalancer.getLoadBalancerArn());
	      
			this.otherElbSheet.addMergedRegion(new CellRangeAddress(row.getRowNum(), row.getRowNum(), 11, 13));
	      
			StringBuffer tsb = new StringBuffer();
	      
			DescribeTagsResult describeTagsResult = this.amazonClients.AmazonElasticLoadBalancing2.describeTags(new DescribeTagsRequest().withResourceArns(loadBalancer.getLoadBalancerArn()));
			List<TagDescription> tagDescriptions = describeTagsResult.getTagDescriptions();
			for(TagDescription tagDescription : tagDescriptions) {
				for(com.amazonaws.services.elasticloadbalancingv2.model.Tag tag : tagDescription.getTags()) {
					if(tsb.length() > 0) tsb.append("\n");
					tsb.append("Key=");
					tsb.append(tag.getKey());
					tsb.append(", Value=");
					tsb.append(tag.getValue());
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
	    	DescribeListenersResult describeListenersResult = this.amazonClients.AmazonElasticLoadBalancing2.describeListeners(new DescribeListenersRequest().withLoadBalancerArn(loadBalancer.getLoadBalancerArn()));
	      
	    	List<com.amazonaws.services.elasticloadbalancingv2.model.Listener> listeners = describeListenersResult.getListeners();
	    	for (com.amazonaws.services.elasticloadbalancingv2.model.Listener listener : listeners) {
	    	
	    		row = this.xssfHelper.createRow(this.otherElbSheet, 1);
	    		this.xssfHelper.setLeftThinCell(row, Integer.toString(row.getRowNum() - 1));
	    		this.xssfHelper.setSubHeadCell(row, listener.getProtocol());
	    		this.xssfHelper.setSubHeadCell(row, listener.getPort().toString());

	    		if("network".equals(loadBalancer.getType())) {
	    			List<Action> actions = listener.getDefaultActions();
	    			if(actions != null) {
	    				this.makeOtherELBTargetGroupActionHeader("Default");
	    				this.makeOtherELBActions(actions);
	    			}
	    		}
	    	
	    		if("application".equals(loadBalancer.getType())) {
	    			DescribeRulesResult describeRulesResult = this.amazonClients.AmazonElasticLoadBalancing2.describeRules(new DescribeRulesRequest().withListenerArn(listener.getListenerArn()));
	    			List<Rule> rules = describeRulesResult.getRules();
	    			for (Rule rule : rules) {
	    				this.makeOtherELBRules(rule);
	    				this.makeOtherELBTargetGroupActionHeader("Rule");
	    				this.makeOtherELBActions(rule.getActions());
	    			}
	    		}
	    	
	    		this.xssfHelper.setSubHeadCell(row, listener.getSslPolicy());
	    	
	    		StringBuffer scer = new StringBuffer();
	    		List<Certificate> certificates = listener.getCertificates();
	    		if(certificates != null) {
	    			for(Certificate certificate : certificates) {
	    				if(scer.length() > 0) scer.append("\n");
	    				scer.append("Certificate ARN=");
	    				scer.append(certificate.getCertificateArn());
	    				scer.append(", isDefault=");
	    				scer.append(certificate.getIsDefault());
	    			}
	    		}
	    		this.xssfHelper.setSubHeadCell(row, scer.toString() + "");
	    		this.xssfHelper.setSubHeadCell(row, "");
	    		this.xssfHelper.setSubHeadCell(row, "");
	    		this.xssfHelper.setSubHeadCell(row, "");
	    		this.xssfHelper.setSubHeadCell(row, listener.getListenerArn());
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
		List<RuleCondition> ruleConditions = rule.getConditions();
		for (RuleCondition ruleCondition : ruleConditions) {
		  
			ruleCs.append(ruleCondition.getField());
			ruleCs.append("=");
		  
			StringBuffer fieldVs = new StringBuffer();
			List<String> fieldValues = ruleCondition.getValues();
			for(String fieldValue : fieldValues) {
				if(fieldVs.length() > 0) fieldVs.append(", ");
				fieldVs.append(fieldValue);  
			}
	      
			ruleCs.append(fieldVs.toString());
		}
	  
		row = this.xssfHelper.createRow(this.otherElbSheet, 1);
		this.xssfHelper.setLeftThinCell(row, Integer.toString(row.getRowNum() - 1));
		this.xssfHelper.setCell(row, rule.getIsDefault().toString());
		this.xssfHelper.setCell(row, rule.getPriority());
		this.xssfHelper.setCell(row, ruleCs.toString());
		this.xssfHelper.setCell(row, "");
		this.xssfHelper.setCell(row, "");
		this.xssfHelper.setCell(row, rule.getRuleArn());
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
	    	DescribeTargetGroupsResult describeTargetGroupsResult = this.amazonClients.AmazonElasticLoadBalancing2.describeTargetGroups(new DescribeTargetGroupsRequest().withTargetGroupArns(action.getTargetGroupArn()));
	    	List<TargetGroup> targetGroups = describeTargetGroupsResult.getTargetGroups();
	    	for(TargetGroup targetGroup : targetGroups) {

  	    	this.makeOtherELBTargetGroupHeader();
  	    	
	    		StringBuffer tga = new StringBuffer();
	    		DescribeTargetGroupAttributesResult describeTargetGroupAttributesResult = this.amazonClients.AmazonElasticLoadBalancing2.describeTargetGroupAttributes(new DescribeTargetGroupAttributesRequest().withTargetGroupArn(targetGroup.getTargetGroupArn()));
	    		List<TargetGroupAttribute> targetGroupAttributes = describeTargetGroupAttributesResult.getAttributes();
	    		for(TargetGroupAttribute targetGroupAttribute : targetGroupAttributes) {
	    			if(tga.length() > 0) tga.append("\n");
	    			tga.append(targetGroupAttribute.getKey());
	    			tga.append("=");
	    			tga.append(targetGroupAttribute.getValue());
	    		}
	    		
	    		this.makeOtherELBTargetGroup(action, targetGroup, tga.toString());	    	    		
	    		
	    		this.makeOtherELBTargetHealthHeader();
	    		DescribeTargetHealthResult describeTargetHealthResult = this.amazonClients.AmazonElasticLoadBalancing2.describeTargetHealth(new DescribeTargetHealthRequest().withTargetGroupArn(targetGroup.getTargetGroupArn()));
	    		List<TargetHealthDescription> targetHealthDescriptions = describeTargetHealthResult.getTargetHealthDescriptions();
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
	  
		TargetDescription targetDescription = targetHealthDescription.getTarget();
		TargetHealth targetHealth = targetHealthDescription.getTargetHealth();
		
		XSSFRow row = this.xssfHelper.createRow(this.otherElbSheet, 1);
		this.xssfHelper.setLeftThinCell(row, Integer.toString(row.getRowNum() - 1));
		this.xssfHelper.setCell(row, targetDescription.getAvailabilityZone());
		this.xssfHelper.setCell(row, targetDescription.getId());
		this.xssfHelper.setCell(row, targetDescription.getPort().toString());
		this.xssfHelper.setCell(row, targetHealthDescription.getHealthCheckPort());
		this.xssfHelper.setCell(row, targetHealth == null ? "" : targetHealth.getState());
		this.xssfHelper.setCell(row, targetHealth == null ? "" : targetHealth.getReason());
		this.xssfHelper.setCell(row, targetHealth == null ? "" : targetHealth.getDescription());
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
		this.xssfHelper.setCell(row, action.getOrder() == null ? "" : action.getOrder().toString());
		this.xssfHelper.setCell(row, action.getType());
		this.xssfHelper.setCell(row, targetGroup.getTargetGroupName());
		this.xssfHelper.setCell(row, targetGroup.getTargetType());
		this.xssfHelper.setCell(row, targetGroup.getProtocol());
		this.xssfHelper.setCell(row, targetGroup.getPort().toString());
		this.xssfHelper.setCell(row, targetGroup.getHealthCheckProtocol());
		this.xssfHelper.setCell(row, targetGroup.getHealthCheckPort());
		this.xssfHelper.setCell(row, targetGroup.getHealthCheckPath());
		this.xssfHelper.setCell(row, targetGroup.getHealthCheckIntervalSeconds().toString());
		this.xssfHelper.setCell(row, targetGroup.getHealthCheckTimeoutSeconds().toString());
		this.xssfHelper.setCell(row, targetGroup.getHealthyThresholdCount().toString());
		this.xssfHelper.setCell(row, targetGroup.getUnhealthyThresholdCount().toString());
		this.xssfHelper.setCell(row, targetGroup.getTargetGroupArn());
		this.xssfHelper.setCell(row, targetGroupAttributes);
	}

	
	private void makeEC2Instance() {
		XSSFRow row = null;
		List<Reservation> reservations = this.amazonClients.AmazonEC2.describeInstances().getReservations();
		for (Reservation reservation : reservations) {
			List<com.amazonaws.services.ec2.model.Instance> instances = reservation.getInstances();
			for (com.amazonaws.services.ec2.model.Instance instance : instances) {
				Subnet subnet = (Subnet) this.amazonClients.AmazonEC2
						.describeSubnets(
								new DescribeSubnetsRequest().withSubnetIds(new String[] { instance.getSubnetId() }))
						.getSubnets().get(0);

				row = this.xssfHelper.createRow(this.ec2InstanceSheet, 1);
				this.xssfHelper.setLeftThinCell(row, Integer.toString(row.getRowNum() - 1));
				this.xssfHelper.setCell(row, getNameTagValue(instance.getTags()));
				this.xssfHelper.setCell(row, instance.getInstanceId());
				this.xssfHelper.setCell(row, instance.getInstanceType());
				this.xssfHelper.setCell(row, subnet.getAvailabilityZone());
				this.xssfHelper.setCell(row, instance.getPrivateIpAddress());
				this.xssfHelper.setCell(row, instance.getPublicIpAddress());
				this.xssfHelper.setCell(row, instance.getPrivateDnsName());
				this.xssfHelper.setCell(row, instance.getPublicDnsName());

				StringBuffer securityGroups = new StringBuffer();
				List<GroupIdentifier> groupIdentifiers = instance.getSecurityGroups();
				for (GroupIdentifier groupIdentifier : groupIdentifiers) {
					if (securityGroups.length() > 0) {
						securityGroups.append("\n");
					}
					securityGroups.append(groupIdentifier.getGroupName());
					securityGroups.append(" (");
					securityGroups.append(groupIdentifier.getGroupId());
					securityGroups.append(")");
				}
				this.xssfHelper.setCell(row, securityGroups.toString());

				Date launchTime = instance.getLaunchTime();
				this.xssfHelper.setCell(row, formatDate.format(launchTime));

				InstanceState instanceState = instance.getState();
				this.xssfHelper.setCell(row,
						instanceState == null ? "" : instanceState.getCode() + "(" + instanceState.getName() + ")");

				StateReason stateReason = instance.getStateReason();
				this.xssfHelper.setCell(row,
						stateReason == null ? "" : stateReason.getCode() + "(" + stateReason.getMessage() + ")");
				this.xssfHelper.setCell(row, instance.getStateTransitionReason());

				Monitoring monitoring = instance.getMonitoring();
				this.xssfHelper.setCell(row, monitoring == null ? "" : monitoring.getState());

				this.xssfHelper.setCell(row, Boolean
						.toString(instance.getSourceDestCheck() == null ? false : instance.getSourceDestCheck()));
				this.xssfHelper.setCell(row, instance.getSpotInstanceRequestId());
				this.xssfHelper.setCell(row, instance.getSriovNetSupport());

				this.xssfHelper.setCell(row, instance.getVirtualizationType());
				this.xssfHelper.setCell(row, instance.getPlatform());
				this.xssfHelper.setCell(row, instance.getArchitecture());
				this.xssfHelper.setCell(row, instance.getKernelId());
				this.xssfHelper.setCell(row,
						instance.getEnaSupport() == null ? "" : Boolean.toString(instance.getEnaSupport()));
				this.xssfHelper.setCell(row, instance.getHypervisor());
				this.xssfHelper.setCell(row, instance.getClientToken());
				this.xssfHelper.setCell(row, Integer.toString(instance.getAmiLaunchIndex()));
				this.xssfHelper.setCell(row, instance.getImageId());
				this.xssfHelper.setCell(row, instance.getInstanceLifecycle());
				this.xssfHelper.setCell(row, instance.getKeyName());

				CpuOptions cpuOptions = instance.getCpuOptions();
				this.xssfHelper.setCell(row,
						cpuOptions == null ? ""
								: "Core : " + cpuOptions.getCoreCount() + "\nThreads Per Core : "
										+ cpuOptions.getThreadsPerCore());

				StringBuffer gpuAssociation = new StringBuffer();
				List<ElasticGpuAssociation> elasticGpuAssociations = instance.getElasticGpuAssociations();
				for (ElasticGpuAssociation elasticGpuAssociation : elasticGpuAssociations) {
					if (gpuAssociation.length() > 0) {
						gpuAssociation.append("\n");
					}
					gpuAssociation.append("AssociationID : " + elasticGpuAssociation.getElasticGpuAssociationId());
					gpuAssociation.append(", GpuID : " + elasticGpuAssociation.getElasticGpuId());
					gpuAssociation
							.append(", AssociationTime : " + elasticGpuAssociation.getElasticGpuAssociationTime());
					gpuAssociation
							.append(", AssociationState : " + elasticGpuAssociation.getElasticGpuAssociationState());
				}
				this.xssfHelper.setCell(row, gpuAssociation.toString());

				this.xssfHelper.setCell(row, Boolean.toString(instance.getEbsOptimized()));
				this.xssfHelper.setCell(row, instance.getRamdiskId());
				this.xssfHelper.setCell(row, instance.getRootDeviceName());
				this.xssfHelper.setCell(row, instance.getRootDeviceType());

				StringBuffer blockDeviceMapping = new StringBuffer();
				List<InstanceBlockDeviceMapping> instanceBlockDeviceMappings = instance.getBlockDeviceMappings();
				for (InstanceBlockDeviceMapping instanceBlockDeviceMapping : instanceBlockDeviceMappings) {
					if (blockDeviceMapping.length() > 0) {
						blockDeviceMapping.append("\n");
					}
					blockDeviceMapping.append("DeviceName : " + instanceBlockDeviceMapping.getDeviceName());
					EbsInstanceBlockDevice ebsInstanceBlockDevice = instanceBlockDeviceMapping.getEbs();
					if (ebsInstanceBlockDevice != null) {
						blockDeviceMapping.append(", VolumeId : " + ebsInstanceBlockDevice.getVolumeId());
						blockDeviceMapping
								.append(", Attach Time : " + formatDate.format(ebsInstanceBlockDevice.getAttachTime()));
						blockDeviceMapping.append(", DeleteOnTermination : "
								+ Boolean.toString(ebsInstanceBlockDevice.getDeleteOnTermination()));
						blockDeviceMapping.append(", Status : " + ebsInstanceBlockDevice.getStatus());
					}
				}
				this.xssfHelper.setCell(row, blockDeviceMapping.toString());

				StringBuffer networkInterfaces = new StringBuffer();
				List<InstanceNetworkInterface> instanceNetworkInterfaces = instance.getNetworkInterfaces();
				for (InstanceNetworkInterface instanceNetworkInterface : instanceNetworkInterfaces) {
					if (networkInterfaces.length() > 0) {
						networkInterfaces.append("\n");
					}
					networkInterfaces
							.append("NetworkInterfaceid : " + instanceNetworkInterface.getNetworkInterfaceId());
					networkInterfaces.append(", OwnerId : " + instanceNetworkInterface.getOwnerId());
					networkInterfaces.append(", VpcId : " + instanceNetworkInterface.getVpcId());
					networkInterfaces.append(", SubnetId : " + instanceNetworkInterface.getSubnetId());
					networkInterfaces.append(", Status : " + instanceNetworkInterface.getStatus());
					networkInterfaces.append(", SourceDestCheck : " + instanceNetworkInterface.getSourceDestCheck());
					networkInterfaces.append(", Mac : " + instanceNetworkInterface.getMacAddress());
					networkInterfaces
							.append(", Private Ip Address : " + instanceNetworkInterface.getPrivateIpAddress());
					networkInterfaces.append(", Private Dns Name : " + instanceNetworkInterface.getPrivateDnsName());

					List<InstancePrivateIpAddress> instancePrivateIpAddresses = instanceNetworkInterface
							.getPrivateIpAddresses();
					networkInterfaces.append(", Private Ips [");
					for (InstancePrivateIpAddress instancePrivateIpAddress : instancePrivateIpAddresses) {
						networkInterfaces.append("[ Primary : " + instancePrivateIpAddress.getPrimary());
						networkInterfaces.append(", Ip : " + instancePrivateIpAddress.getPrivateIpAddress());
						networkInterfaces.append(", Dns Name : " + instancePrivateIpAddress.getPrivateDnsName());
						networkInterfaces.append("], ");
					}
					networkInterfaces.append("]");

					List<InstanceIpv6Address> instanceIpv6Addresses = instanceNetworkInterface.getIpv6Addresses();
					networkInterfaces.append(", Ipv6s [");
					for (InstanceIpv6Address instanceIpv6Address : instanceIpv6Addresses) {
						networkInterfaces.append(instanceIpv6Address.getIpv6Address() + ", ");
					}
					networkInterfaces.append("]");

					InstanceNetworkInterfaceAssociation instanceNetworkInterfaceAssociation = instanceNetworkInterface
							.getAssociation();
					if (instanceNetworkInterfaceAssociation != null) {
						networkInterfaces.append("Association [");
						networkInterfaces.append("IpOwnerId : " + instanceNetworkInterfaceAssociation.getIpOwnerId());
						networkInterfaces.append(", PublicIP : " + instanceNetworkInterfaceAssociation.getPublicIp());
						networkInterfaces
								.append(", PublicDnsName : " + instanceNetworkInterfaceAssociation.getPublicDnsName());
						networkInterfaces.append("]");
					}

					InstanceNetworkInterfaceAttachment instanceNetworkInterfaceAttachment = instanceNetworkInterface
							.getAttachment();
					if (instanceNetworkInterfaceAttachment != null) {
						networkInterfaces.append(", Attachment [");
						networkInterfaces
								.append("AttachmentId : " + instanceNetworkInterfaceAttachment.getAttachmentId());
						networkInterfaces.append(", AttachTime : "
								+ formatDate.format(instanceNetworkInterfaceAttachment.getAttachTime()));
						networkInterfaces.append(", DeleteOnTermination : "
								+ Boolean.toString(instanceNetworkInterfaceAttachment.getDeleteOnTermination()));
						networkInterfaces.append(", DeviceIndex : "
								+ Integer.toString(instanceNetworkInterfaceAttachment.getDeviceIndex()));
						networkInterfaces.append(", Status : " + instanceNetworkInterfaceAttachment.getStatus());
						networkInterfaces.append("]");
					}

					networkInterfaces.append(", Description : " + instanceNetworkInterface.getDescription());
					List<GroupIdentifier> ngroupIdentifiers = instanceNetworkInterface.getGroups();
					networkInterfaces.append(", Groups [");
					for (GroupIdentifier ngroupIdentifier : ngroupIdentifiers) {
						networkInterfaces
								.append(ngroupIdentifier.getGroupId() + "(" + ngroupIdentifier.getGroupName() + ")");
						networkInterfaces.append(", ");
					}
					networkInterfaces.append("]");

				}
				this.xssfHelper.setCell(row, networkInterfaces.toString());

				Placement placement = instance.getPlacement();
				this.xssfHelper.setCell(row,
						placement == null ? ""
								: "Affinify : " + placement.getAffinity() + "\nAvailabilityZone : "
										+ placement.getAvailabilityZone() + "\nGroupName : " + placement.getGroupName()
										+ "\nHostId : " + placement.getHostId() + "\nSpreadDomain : "
										+ placement.getSpreadDomain() + "\nTenancy : " + placement.getTenancy());

				StringBuffer productCds = new StringBuffer();
				List<ProductCode> productCodes = instance.getProductCodes();
				for (ProductCode productCode : productCodes) {
					if (productCds.length() > 0) {
						productCds.append("\n");
					}
					productCds.append("ID : " + productCode.getProductCodeId());
					productCds.append(", TYPE : " + productCode.getProductCodeType());
				}
				this.xssfHelper.setCell(row, productCds.toString());

				IamInstanceProfile iamInstanceProfile = instance.getIamInstanceProfile();
				this.xssfHelper.setCell(row, iamInstanceProfile == null ? ""
						: "ID : " + iamInstanceProfile.getId() + ", ARN : " + iamInstanceProfile.getArn());

				this.xssfHelper.setRightThinCell(row, this.getAllTagValue(instance.getTags()));
			}
		}

	}

	private void makeSecurityGroup() {
		XSSFRow row = null;

		int securityGroupStartRowNum = this.securityGroupSheet.getLastRowNum();

		List<SecurityGroup> securityGroups = this.amazonClients.AmazonEC2.describeSecurityGroups().getSecurityGroups();
		for (SecurityGroup securityGroup : securityGroups) {
			row = this.xssfHelper.createRow(this.securityGroupSheet, 1);
			this.xssfHelper.setLeftThinCell(row, Integer.toString(row.getRowNum() - 1));
			this.xssfHelper.setCell(row, this.region.getName());
			this.xssfHelper.setCell(row, securityGroup.getVpcId());
			this.xssfHelper.setBrownBoldLeftThinCell(row, securityGroup.getGroupId() + "["
					+ securityGroup.getGroupName() + "] - " + securityGroup.getDescription());
			this.xssfHelper.setCell(row, "");
			this.xssfHelper.setCell(row, "");
			this.xssfHelper.setCell(row, "");
			this.xssfHelper.setCell(row, "");
			this.xssfHelper.setBrownBoldLeftThinCell(row, this.getAllTagValue(securityGroup.getTags()));
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

			List<IpPermission> inbounds = securityGroup.getIpPermissions();
			List<IpPermission> outbounds = securityGroup.getIpPermissionsEgress();

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
			IpPermissionsSize += ipPermission.getPrefixListIds().size();
			IpPermissionsSize += ipPermission.getIpv4Ranges().size();
			IpPermissionsSize += ipPermission.getIpv6Ranges().size();
			IpPermissionsSize += ipPermission.getUserIdGroupPairs().size();
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

		this.xssfHelper.setCell(row, 3 + (isEngress ? 5 : 0), getSecurityGroupRuleType(ipPermission.getIpProtocol(),
				ipPermission.getFromPort(), ipPermission.getToPort()));
		this.xssfHelper.setCell(row, 4 + (isEngress ? 5 : 0),
				("-1".equals(ipPermission.getIpProtocol())) || (("icmp".equals(ipPermission.getIpProtocol()))
						&& (ipPermission.getFromPort() != null) && (ipPermission.getFromPort().intValue() == -1))
								? "All"
								: ipPermission.getIpProtocol().toUpperCase());
		this.xssfHelper.setCell(row, 5 + (isEngress ? 5 : 0),
				getPortRange(ipPermission.getFromPort(), ipPermission.getToPort()));
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

			List<PrefixListId> prefixListIds = ipPermission.getPrefixListIds();
			for (PrefixListId prefixListId : prefixListIds) {
				row = getIpPermissionCom(ipPermission, startRowNum + rowNum, prefixListId.getPrefixListId(),
						prefixListId.getDescription(), isEngress);
				rowNum++;
			}
			List<IpRange> ipRanges = ipPermission.getIpv4Ranges();
			for (IpRange ipRange : ipRanges) {
				row = getIpPermissionCom(ipPermission, startRowNum + rowNum, ipRange.getCidrIp(),
						ipRange.getDescription(), isEngress);
				rowNum++;
			}
			List<Ipv6Range> ipv6Ranges = ipPermission.getIpv6Ranges();
			for (Ipv6Range ipv6Range : ipv6Ranges) {
				row = getIpPermissionCom(ipPermission, startRowNum + rowNum, ipv6Range.getCidrIpv6(),
						ipv6Range.getDescription(), isEngress);
				rowNum++;
			}
			List<UserIdGroupPair> userIdGroupPairs = ipPermission.getUserIdGroupPairs();
			for (UserIdGroupPair userIdGroupPair : userIdGroupPairs) {
				row = getIpPermissionCom(ipPermission, startRowNum + rowNum,
						getUserIdGroupPair(userIdGroupPair.getGroupId()), userIdGroupPair.getDescription(), isEngress);
				rowNum++;
			}
		}
		return row;
	}

	private String getSGName(String GroupId) {
		String sgName = "";
		try {
			DescribeSecurityGroupsResult describeSecurityGroupsResult = this.amazonClients.AmazonEC2
					.describeSecurityGroups(new DescribeSecurityGroupsRequest().withGroupIds(new String[] { GroupId }));
			sgName = ((SecurityGroup) describeSecurityGroupsResult.getSecurityGroups().get(0)).getGroupName();
		} catch (Exception localException) {
		}
		return sgName;
	}

	
	private void makeInternetGateway() {
		XSSFRow row = null;
		List<InternetGateway> internetGateways = this.amazonClients.AmazonEC2
				.describeInternetGateways(new DescribeInternetGatewaysRequest()).getInternetGateways();
		for (int iIGW = 0; iIGW < internetGateways.size(); iIGW++) {
			InternetGateway internetGateway = (InternetGateway) internetGateways.get(iIGW);

			row = this.xssfHelper.createRow(this.internetGatewaySheet, 1);
			this.xssfHelper.setLeftThinCell(row, Integer.toString(row.getRowNum() - 1));
			this.xssfHelper.setCell(row, this.region.getName());

			StringBuffer vpcBuffer = new StringBuffer();
			List<InternetGatewayAttachment> internetGatewayAttachments = internetGateway.getAttachments();
			for (InternetGatewayAttachment internetGatewayAttachment : internetGatewayAttachments) {
				if (vpcBuffer.length() > 0) {
					vpcBuffer.append("\n");
				}
				vpcBuffer.append(internetGatewayAttachment.getVpcId());
				vpcBuffer.append("(");
				vpcBuffer.append(internetGatewayAttachment.getState());
				vpcBuffer.append(")");
			}
			this.xssfHelper.setCell(row, vpcBuffer.toString());
			this.xssfHelper.setCell(row, getNameTagValue(internetGateway.getTags()));
			this.xssfHelper.setCell(row, internetGateway.getInternetGatewayId());
			this.xssfHelper.setRightThinCell(row, this.getAllTagValue(internetGateway.getTags()));
		}
	}

	
	private HashMap<String, String> makeRouteTable() {
		XSSFRow row = null;

		HashMap<String, String> vpcMainRouteTables = new HashMap<String, String>();
		List<RouteTable> routeTables = this.amazonClients.AmazonEC2.describeRouteTables().getRouteTables();
		for (int i = 0; i < routeTables.size(); i++) {
			boolean isMain = false;
			RouteTable routeTable = (RouteTable) routeTables.get(i);

			List<RouteTableAssociation> routeTableAssociations = routeTable.getAssociations();
			for (RouteTableAssociation routeTableAssociation : routeTableAssociations) {
				if (routeTableAssociation.isMain().booleanValue()) {
					isMain = true;
					vpcMainRouteTables.put(routeTable.getVpcId(),
							routeTable.getRouteTableId() + " | " + getNameTagValue(routeTable.getTags()));
				}
			}
			row = this.xssfHelper.createRow(this.routeTableSheet, 1);
			int firstRowNum = row.getRowNum();
			this.xssfHelper.setLeftThinCell(row, Integer.toString(row.getRowNum() - 1));
			this.xssfHelper.setCell(row, this.region.getName());
			this.xssfHelper.setCell(row, routeTable.getVpcId());
			this.xssfHelper.setBrownBoldLeftThinCell(row, getNameTagValue(routeTable.getTags()));
			this.xssfHelper.setCell(row, routeTable.getRouteTableId());
			this.xssfHelper.setLeftRightThinCell(row, isMain ? "YES" : "NO");
			this.xssfHelper.setRightThinCell(row, "");
			this.xssfHelper.setCell(row, this.getAllTagValue(routeTable.getTags()));
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

			List<Route> routes = routeTable.getRoutes();
			for (int iRoute = 0; iRoute < routes.size(); iRoute++) {
				Route route = routes.get(iRoute);

				String destination = "";
				String target = "";
				if (route.getDestinationCidrBlock() != null) {
					destination = route.getDestinationCidrBlock();
				} else if (route.getDestinationIpv6CidrBlock() != null) {
					destination = route.getDestinationIpv6CidrBlock();
				} else if (route.getDestinationPrefixListId() != null) {
					destination = route.getDestinationPrefixListId();
				}
				if (route.getEgressOnlyInternetGatewayId() != null) {
					target = route.getEgressOnlyInternetGatewayId();
				} else if (route.getGatewayId() != null) {
					target = route.getGatewayId();
				} else if (route.getNatGatewayId() != null) {
					target = route.getNatGatewayId();
				} else if (route.getNetworkInterfaceId() != null) {
					target = route.getNetworkInterfaceId();
				} else if (route.getVpcPeeringConnectionId() != null) {
					target = route.getVpcPeeringConnectionId();
				}
				row = this.xssfHelper.createRow(this.routeTableSheet, 1);
				this.xssfHelper.setLeftThinCell(row, "");
				this.xssfHelper.setCell(row, "");
				this.xssfHelper.setCell(row, "");
				this.xssfHelper.setLeftThinCell(row, destination);
				this.xssfHelper.setCell(row, target);
				this.xssfHelper.setCell(row, route.getState());
				this.xssfHelper.setCell(row, route.getOrigin());
				this.xssfHelper.setLeftThinCell(row, "");
				this.xssfHelper.setRightThinCell(row, "");

				rows.add(Integer.valueOf(row.getRowNum()));
			}
			int lastRowNum = row.getRowNum();

			int iSubnetCnt = 0;
			for (int iAssoc = 0; iAssoc < routeTableAssociations.size(); iAssoc++) {
				RouteTableAssociation routeTableAssociation = routeTableAssociations.get(iAssoc);

				if (!routeTableAssociation.getMain()) {
					List<Subnet> subnets = this.amazonClients.AmazonEC2.describeSubnets(new DescribeSubnetsRequest()
							.withSubnetIds(new String[] { routeTableAssociation.getSubnetId() })).getSubnets();
					for (int iSubnet = 0; iSubnet < subnets.size(); iSubnet++) {
						Subnet subnet = (Subnet) subnets.get(iSubnet);
						if (rows.size() > iSubnetCnt) {
							row = this.routeTableSheet.getRow(((Integer) rows.get(iSubnetCnt)).intValue());
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
									subnet.getSubnetId() + " | " + getNameTagValue(subnet.getTags()));
							this.xssfHelper.setRightThinCell(row, subnet.getCidrBlock());
						} else {
							row.getCell(row.getLastCellNum() - 2)
									.setCellValue(subnet.getSubnetId() + " | " + getNameTagValue(subnet.getTags()));
							row.getCell(row.getLastCellNum() - 1).setCellValue(subnet.getCidrBlock());
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

		List<Subnet> subnets = this.amazonClients.AmazonEC2.describeSubnets().getSubnets();
		for (int iSubnet = 0; iSubnet < subnets.size(); iSubnet++) {
			Subnet subnet = (Subnet) subnets.get(iSubnet);

			List<String> subnetList = new ArrayList<String>();
			subnetList.add(subnet.getSubnetId());

			List<RouteTable> routeTables = this.amazonClients.AmazonEC2
					.describeRouteTables(new DescribeRouteTablesRequest()
							.withFilters(new Filter[] { new Filter("association.subnet-id", subnetList) }))
					.getRouteTables();
			List<NetworkAcl> networkAcls = this.amazonClients.AmazonEC2
					.describeNetworkAcls(new DescribeNetworkAclsRequest()
							.withFilters(new Filter[] { new Filter("association.subnet-id", subnetList) }))
					.getNetworkAcls();

			String routeTableInfo = "";
			String networkAclInfo = "";
			if (routeTables.size() > 0) {
				RouteTable routeTable = (RouteTable) routeTables.get(0);
				routeTableInfo = routeTable.getRouteTableId() + " | " + getNameTagValue(routeTable.getTags());
			}
			if (networkAcls.size() > 0) {
				NetworkAcl networkAcl = networkAcls.get(0);
				networkAclInfo = networkAcl.getNetworkAclId();
			}
			row = this.xssfHelper.createRow(this.subnetSheet, 1);
			this.xssfHelper.setLeftThinCell(row, Integer.toString(row.getRowNum() - 1));
			this.xssfHelper.setCell(row, this.region.getName());
			this.xssfHelper.setCell(row, subnet.getVpcId());
			this.xssfHelper.setCell(row, getNameTagValue(subnet.getTags()));
			this.xssfHelper.setCell(row, subnet.getSubnetId());
			this.xssfHelper.setCell(row, subnet.getCidrBlock());
			this.xssfHelper.setCell(row, subnet.getAvailabilityZone());
			this.xssfHelper.setCell(row, routeTableInfo);
			this.xssfHelper.setCell(row, networkAclInfo);

			this.xssfHelper.setCell(row, Boolean.toString(subnet.getAssignIpv6AddressOnCreation()));
			this.xssfHelper.setCell(row, Integer.toString(subnet.getAvailableIpAddressCount()));
			this.xssfHelper.setCell(row, Boolean.toString(subnet.getDefaultForAz()));

			StringBuffer subnetIpv6Cidr = new StringBuffer();
			List<SubnetIpv6CidrBlockAssociation> subnetIpv6CidrBlockAssociations = subnet
					.getIpv6CidrBlockAssociationSet();
			for (SubnetIpv6CidrBlockAssociation subnetIpv6CidrBlockAssociation : subnetIpv6CidrBlockAssociations) {
				subnetIpv6Cidr.append("AssociationId : " + subnetIpv6CidrBlockAssociation.getAssociationId());
				subnetIpv6Cidr.append("\n");
				subnetIpv6Cidr.append("CIDR Block : " + subnetIpv6CidrBlockAssociation.getIpv6CidrBlock());
				subnetIpv6Cidr.append("\n");
				SubnetCidrBlockState subnetCidrBlockState = subnetIpv6CidrBlockAssociation.getIpv6CidrBlockState();
				if (subnetCidrBlockState != null) {
					subnetIpv6Cidr.append("State : " + subnetCidrBlockState.getState());
					subnetIpv6Cidr.append("\n");
					subnetIpv6Cidr.append("Status Message : " + subnetCidrBlockState.getStatusMessage());
				}
			}
			this.xssfHelper.setCell(row, subnetIpv6Cidr.toString());
			this.xssfHelper.setCell(row, Boolean.toString(subnet.getMapPublicIpOnLaunch()));
			this.xssfHelper.setCell(row, subnet.getState());
			this.xssfHelper.setCell(row, this.getAllTagValue(subnet.getTags()));

			this.xssfHelper.setRightThinCell(row, "");

		}
	}

}
