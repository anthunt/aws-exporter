package anthunt.aws.exporter;

import com.amazonaws.regions.Regions;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.net.URISyntaxException;
import java.util.ArrayList;
import java.util.Enumeration;
import java.util.HashMap;
import java.util.List;
import java.util.Map;
import java.util.Properties;
import java.util.Scanner;
import java.util.Set;

import org.apache.maven.model.building.ModelBuildingException;
import org.eclipse.aether.resolution.DependencyResolutionException;

public class AWSExportStarter
{
  private static final String dataFileName = "conf/awsExporter.data";
  private Scanner scanner;
  private Properties properties;
  private File propertiesFile;
  private HashMap<String, AmazonAccess> amazonAccesses;
  private AmazonAccess amazonAccess;
  
  private static enum PropType
  {
       ACCESS_KEY(".aws.accessKey")
    ,  SECRET_KEY(".aws.secretKey")
    ,  USE_PROXY(".connect.useProxy")
    ,  PROXY_HOST(".connect.proxy.host")
    ,  PROXY_PORT(".connect.proxy.port")
    ,  CROSS_ACCOUNT(".aws.cross")
    ;
	  
    private String postFix;
    
    private PropType(String postFix)
    {
      this.postFix = postFix;
    }
    
    protected String getPostFix()
    {
      return this.postFix;
    }
    
    protected static PropType getPropType(String key)
    {
      PropType[] propTypes = values();
      for (PropType propType : propTypes) {
        if (key.endsWith(propType.getPostFix())) {
          return propType;
        } else if(key.indexOf(propType.getPostFix()) > 0) {
          return propType;
        }
      }
      return null;
    }
  }
  
  private AWSExportStarter(Scanner scanner)
  {
    this.scanner = scanner;
    this.amazonAccesses = new HashMap<>();
    this.properties = new Properties();
  }
  
  private void execute()
    throws FileNotFoundException, IOException
  {
    loadProperties();
    showAccessList();
    if (this.amazonAccess == null) {
      makeAccess();
    }
  }
  
  private void loadProperties()
    throws FileNotFoundException, IOException
  {
    this.propertiesFile = new File(dataFileName);
    if (this.propertiesFile.exists())
    {
      this.properties.load(new FileInputStream(this.propertiesFile));
      Enumeration<Object> keys = this.properties.keys();
      while (keys.hasMoreElements())
      {
        String key = (String)keys.nextElement();
        setAmazonAccess(key, this.properties.getProperty(key));
      }
    }
    else
    {
      this.propertiesFile.createNewFile();
    }
  }
  
  private void setAmazonAccess(String key, String value)
  {
    PropType propType = PropType.getPropType(key);
    if (propType != null)
    {
      AmazonAccess amazonAccess = null;
      
      String accessType = key.substring(0, key.indexOf(propType.getPostFix()));
      if (this.amazonAccesses.containsKey(accessType)) {
        amazonAccess = (AmazonAccess)this.amazonAccesses.get(accessType);
      } else {
        amazonAccess = new AmazonAccess(accessType);
      }
      switch (propType)
      {
      case ACCESS_KEY: 
        amazonAccess.setAccessKey(value);
        break;
      case SECRET_KEY: 
        amazonAccess.setSecretKey(value);
        break;
      case USE_PROXY: 
        amazonAccess.setUseProxy(Boolean.valueOf(Boolean.parseBoolean(value)));
        break;
      case PROXY_HOST: 
        amazonAccess.setProxyHost(value);
        break;
      case PROXY_PORT: 
        amazonAccess.setProxyPort(Integer.valueOf(Integer.parseInt(value)));
        break;
      case CROSS_ACCOUNT:
    	  
    	String crossAccountKey = key.substring(key.indexOf(propType.getPostFix()) + propType.getPostFix().length() + 1, key.length());
    	
  		String[] crossAccountProps = crossAccountKey.split("[.]");
  		String crossAccountRoleName = crossAccountProps[0];
  		String crossAccountRoleAttr = crossAccountProps[1];
  		
  		CrossAccountRole crossAccountRole = amazonAccess.getCrossAccountRole(crossAccountRoleName);
  		
  		if("accountId".equals(crossAccountRoleAttr)) {
  			crossAccountRole.setCrossAccountId(value);
  			crossAccountRole.setCrossRoleSessionName(crossAccountKey);
  		} else if("roleName".equals(crossAccountRoleAttr)) {
  			crossAccountRole.setCrossRoleName(value);
  		} else if("externId".equals(crossAccountRoleAttr)) {
  			crossAccountRole.setExternId(value);
  		}
  		
      	break;
      }
      
      this.amazonAccesses.put(accessType, amazonAccess);
      
    }
  }
  
  private void showAccessList()
  {
    if (this.amazonAccesses.size() > 0)
    {
      List<String> keyIndex = new ArrayList<>();
      Set<String> keys = this.amazonAccesses.keySet();
      
      System.out.println("접속 번호를 입력하여 주십시요.\n");
      System.out.println("==============================");
      System.out.println("0. 신규  접속 정보 등록");
      System.out.println("------------------------------");
      for (String key : keys)
      {
        keyIndex.add(key);
        System.out.print(keyIndex.size() + ". " + key + " 계정 접속");
        int crossAccountSize = this.amazonAccesses.get(key).getCrossAccountRoles().size();
        if(crossAccountSize > 0) {
        	System.out.print(" - " + crossAccountSize + "개의 CrossAccount Role 사용 가능.");
        }
        System.out.println("");
      }
      System.out.println("------------------------------");
      System.out.println("90. CrossAccount Role 사용");
      System.out.println("------------------------------");
      System.out.println("99. 종료");
      System.out.println("==============================");
      System.out.println("");
      
      checkSelectAccessList(keyIndex);
    }
  }
  
  private void exit() {
	  System.exit(-1);
  }
  
  private void checkSelectAccessList(List<String> keyIndex)
  {
    System.out.print("접속 번호 : ");
    while (!this.scanner.hasNextInt())
    {
      System.out.println("\n잘못된 접속 번호를 선택하셨습니다. 다시 입력해 주십시요.");
      System.out.print("접속 번호 : ");
      this.scanner.nextLine();
    }
    int index = this.scanner.nextInt();
    this.scanner.nextLine();
    if (index == 99)
    {
      this.exit();
    }
    else if (index == 90)
    {
      useCrossRoleAccess(keyIndex);
    }
    else if ((index < 0) || (index > keyIndex.size()))
    {
      System.out.println("잘못된 접속 번호를 선택하셨습니다. 다시 입력해 주십시요.");
      checkSelectAccessList(keyIndex);
    }
    else if (index == 0)
    {
      makeAccess();
    }
    else
    {
      String accessType = (String)keyIndex.get(index - 1);
      this.amazonAccess = this.getAmazonAccess(accessType);
      this.extractRegion(accessType, false, null);      
    }
  }
  
  private void extractRegion(String accessType, boolean isCrossAccount, String crossAccountKey) {
	  List<Regions> regions = new AWSRegionSelector(this).getRegions();
      
      System.out.println("");
      System.out.println(accessType + " AWS 정보 생성 시작");
      new AWSExporter(this.amazonAccess, regions, isCrossAccount, crossAccountKey);
      System.out.println(accessType + " AWS 정보 생성 완료");
      System.out.println("");
      
      showAccessList();
  }
  
  private AmazonAccess getAmazonAccess(String accessType) {
	  return (AmazonAccess)this.amazonAccesses.get(accessType);
  }
  
  private void makeAccess()
  {
    System.out.println("\n접속 정보를 입력하여 주십시요.");
    
    String accessType = setAccessType();
    String accessKey = setAccessKey();
    String secretKey = setSecretKey();
    
    AmazonAccess amazonAccess = new AmazonAccess(accessType);
    
    amazonAccess.setAccessKey(accessKey);
    amazonAccess.setSecretKey(secretKey);
    
    boolean isUseProxy = setUseProxy();
    amazonAccess.setUseProxy(Boolean.valueOf(isUseProxy));
    if (isUseProxy)
    {
      String proxyHost = setProxyHost();
      int proxyPort = setProxyPort();
      amazonAccess.setProxyHost(proxyHost);
      amazonAccess.setProxyPort(Integer.valueOf(proxyPort));
    }
    this.amazonAccesses.put(amazonAccess.getAccessType(), amazonAccess);
    
    this.properties.put(amazonAccess.getAccessType() + PropType.ACCESS_KEY.getPostFix(), amazonAccess.getAccessKey());
    this.properties.put(amazonAccess.getAccessType() + PropType.SECRET_KEY.getPostFix(), amazonAccess.getSecretKey());
    this.properties.put(amazonAccess.getAccessType() + PropType.USE_PROXY.getPostFix(), amazonAccess.isUseProxy().toString());
    
    if(amazonAccess.isUseProxy()) {
	    this.properties.put(amazonAccess.getAccessType() + PropType.PROXY_HOST.getPostFix(), amazonAccess.getProxyHost());
	    this.properties.put(amazonAccess.getAccessType() + PropType.PROXY_PORT.getPostFix(), amazonAccess.getProxyPort().toString());
    }
    
    storeProperties();
    try
    {
      execute();
    }
    catch (FileNotFoundException e)
    {
      e.printStackTrace();
    }
    catch (IOException e)
    {
      e.printStackTrace();
    }
  }
  
  private void useCrossRoleAccess(List<String> keyIndex) {
	  
	  System.out.println("\n\nCrossAccount Role 접속을 위한 계정 접속번호를 입력하여 주십시요.\n");
      System.out.println("==============================");
	  for (int i = 0; i < keyIndex.size(); i++) {
		  String key = keyIndex.get(i);
		  System.out.println((i + 1) + ". " + key + " 계정 접속");
	  }
	  System.out.println("------------------------------\n");
	  System.out.println("90. 전 단계로 돌아가기");
	  System.out.println("99. 종료");
	  System.out.println("==============================\n");
	  
	  System.out.print("CrossAccount Role 접속 번호 : ");
	  while (!this.scanner.hasNextInt())
	  {
		  System.out.println("\n잘못된 접속 번호를 선택하셨습니다. 다시 입력해 주십시요.");
		  System.out.print("CrossAccount Role 접속 번호 : ");
		  this.scanner.nextLine();
	  }
	  int index = this.scanner.nextInt();
	  this.scanner.nextLine();
	  
	  if(index == 99) {
		  this.exit();
	  } else if(index == 90) {
		  this.showAccessList();
	  } else if ((index < 1) || (index > keyIndex.size()))
	  {
	      System.out.println("잘못된 접속 번호를 선택하셨습니다. 다시 입력해 주십시요.");
	      useCrossRoleAccess(keyIndex);
	  } else {
		  showCrossRoleAccess(keyIndex, keyIndex.get(index - 1));
	  }
  }
  
  private void showCrossRoleAccess(List<String> keyIndex, String key) {
	  System.out.println("\n\nCrossAccount Role 접속번호를 입력하여 주십시요.\n");
      System.out.println("==============================");
      System.out.println("0. 신규 CrossAccount Role 등록");
      
      this.amazonAccess = this.amazonAccesses.get(key);
      
      Map<String, CrossAccountRole> crossAccountRoles = this.amazonAccess.getCrossAccountRoles();
      
      if(crossAccountRoles.size() > 0) {
    	  System.out.println("------------------------------\n");
      }
      
      List<String> crossKeyIndex = new ArrayList<>();
      Set<String> keys = crossAccountRoles.keySet();
      for (String crosskey : keys) {
    	  crossKeyIndex.add(crosskey);
    	  System.out.println(crossKeyIndex.size() + ". " + crosskey + " Role 접속");
      }
      
	  System.out.println("------------------------------\n");
	  System.out.println("90. 전 단계로 돌아가기");
	  System.out.println("99. 종료");
	  System.out.println("==============================\n");
	  
	  System.out.print("CrossAccount Role 접속 번호 : ");
	  while (!this.scanner.hasNextInt())
	  {
		  System.out.println("\n잘못된 접속 번호를 선택하셨습니다. 다시 입력해 주십시요.");
		  System.out.print("CrossAccount Role 접속 번호 : ");
		  this.scanner.nextLine();
	  }
	  int index = this.scanner.nextInt();
	  this.scanner.nextLine();
	  
	  if(index == 99) {
		  this.exit();
	  } else if(index == 90) {
		  this.useCrossRoleAccess(keyIndex);
	  } else if(index == 0) {
		  makeCrossRoleAccess(keyIndex, key);
	  } else if ((index < 1) || (index > keyIndex.size()))
	  {
	      System.out.println("잘못된 접속 번호를 선택하셨습니다. 다시 입력해 주십시요.");
	      showCrossRoleAccess(keyIndex, key);
	  } else {
		  extractRegion(key, true, crossKeyIndex.get(index-1));
	  }
	  
  }
  
  private void makeCrossRoleAccess(List<String> keyIndex, String key) {
	  
	  System.out.println("\nCrossAccount Role 접속 정보를 입력하여 주십시요.");
	  
	  String crossAccountKey = this.setCrossAccountKey(key);
	  String crossAccountId = this.setAccountId(key, crossAccountKey);
	  String crossAccountRoleName = this.setAccountRoleName(key, crossAccountKey);
	  String externId = this.setExternId(key, crossAccountKey);
	  
	  AmazonAccess amazonAccess = this.amazonAccesses.get(key);
	  
	  CrossAccountRole crossAccountRole = amazonAccess.getCrossAccountRole(crossAccountKey);
	  crossAccountRole.setCrossAccountId(crossAccountId);
	  crossAccountRole.setCrossRoleName(crossAccountRoleName);
	  crossAccountRole.setCrossRoleSessionName(crossAccountKey);
	  crossAccountRole.setExternId(externId);
	  
	  this.properties.put(amazonAccess.getAccessType() + PropType.CROSS_ACCOUNT.getPostFix() + "." + crossAccountKey + ".accountId", crossAccountRole.getCrossAccountId());
	  this.properties.put(amazonAccess.getAccessType() + PropType.CROSS_ACCOUNT.getPostFix() + "." + crossAccountKey + ".roleName", crossAccountRole.getCrossRoleName());
	  if(crossAccountRole.getExternId() != null) {
		  this.properties.put(amazonAccess.getAccessType() + PropType.CROSS_ACCOUNT.getPostFix() + "." + crossAccountKey + ".externId", crossAccountRole.getExternId());
	  }
	  
	  storeProperties();
	  try {
		loadProperties();
	} catch (FileNotFoundException e) {
		e.printStackTrace();
	} catch (IOException e) {
		e.printStackTrace();
	}
	  
	  showCrossRoleAccess(keyIndex, key);
  }

  private String setCrossAccountKey(String key)
  {
    System.out.print("CrossAccount 접속명(영문) : ");
    while (!this.scanner.hasNextLine()) {
      this.scanner.nextLine();
    }
    String crossAccountKey = this.scanner.nextLine();
    if ((crossAccountKey == null) || ("".equals(crossAccountKey.trim())))
    {
      System.out.println("잘못 입력하셨습니다. CrossAccount 접속명(영문)을 입력하여 주십시요.");
      crossAccountKey = setCrossAccountKey(key);
    }
    return crossAccountKey;
  }
  
  private String setAccountId(String key, String crossAccountKey)
  {
    System.out.print("Account ID : ");
    while (!this.scanner.hasNextLine()) {
      this.scanner.nextLine();
    }
    String accountId = this.scanner.nextLine();
    if ((accountId == null) || ("".equals(accountId.trim())))
    {
      System.out.println("잘못 입력하셨습니다. Account ID를 입력하여 주십시요.");
      accountId = setAccountId(key, crossAccountKey);
    }
    return accountId;
  }
  
  private String setAccountRoleName(String key, String crossAccountKey)
  {
    System.out.print("CrossAccount Role Name : ");
    while (!this.scanner.hasNextLine()) {
      this.scanner.nextLine();
    }
    String accountRoleName = this.scanner.nextLine();
    if ((accountRoleName == null) || ("".equals(accountRoleName.trim())))
    {
      System.out.println("잘못 입력하셨습니다. CrossAccount Role Name을 입력하여 주십시요.");
      accountRoleName = setAccountRoleName(key, crossAccountKey);
    }
    return accountRoleName;
  }
  
  private String setExternId(String key, String crossAccountKey)
  {
    System.out.print("Extern ID : ");
    while (!this.scanner.hasNextLine()) {
      this.scanner.nextLine();
    }
    String externId = this.scanner.nextLine();
    
    return externId == null ? "" : externId;
  }
  
  public void storeProperties()
  {
    try
    {
      this.properties.store(new FileOutputStream(this.propertiesFile), "AWS Exporter Runtime Properties");
    }
    catch (FileNotFoundException e)
    {
      e.printStackTrace();
    }
    catch (IOException e)
    {
      e.printStackTrace();
    }
  }
  
  private String setAccessType()
  {
    System.out.print("접속 명(영문) : ");
    while (!this.scanner.hasNextLine()) {
      this.scanner.nextLine();
    }
    String accessType = this.scanner.nextLine();
    if ((accessType == null) || ("".equals(accessType.trim())))
    {
      System.out.println("잘못 입력하셨습니다. 접속 명(영문)을 입력하여 주십시요.");
      accessType = setAccessType();
    }
    return accessType;
  }
  
  private String setAccessKey()
  {
    System.out.print("AWS AccessKey : ");
    while (!this.scanner.hasNextLine()) {
      this.scanner.nextLine();
    }
    String accessKey = this.scanner.nextLine();
    if ((accessKey == null) || ("".equals(accessKey.trim())))
    {
      System.out.println("잘못 입력하셨습니다. AWS Access Key를 입력하여 주십시요.");
      accessKey = setAccessKey();
    }
    return accessKey;
  }
  
  private String setSecretKey()
  {
    System.out.print("AWS SecretKey : ");
    while (!this.scanner.hasNextLine()) {
      this.scanner.next();
    }
    String secretKey = this.scanner.nextLine();
    if ((secretKey == null) || ("".equals(secretKey.trim())))
    {
      System.out.println("잘못 입력하셨습니다. AWS SecretKey를 입력하여 주십시요.");
      secretKey = setSecretKey();
    }
    return secretKey;
  }
  
  private boolean setUseProxy()
  {
    boolean isUseProxy = false;
    
    System.out.print("Proxy 사용 여부(Y/N) : ");
    while (!this.scanner.hasNextLine()) {
      this.scanner.next();
    }
    String useYN = this.scanner.nextLine();
    if ("Y".equals(useYN))
    {
      isUseProxy = true;
    }
    else if ("N".equals(useYN))
    {
      isUseProxy = false;
    }
    else
    {
      System.out.println("잘못 입력하셨습니다. Proxy 사용 여부(Y/N)를 입력하여 주십시요.");
      isUseProxy = setUseProxy();
    }
    return isUseProxy;
  }
  
  private String setProxyHost()
  {
    System.out.print("Proxy Host : ");
    while (!this.scanner.hasNextLine()) {
      this.scanner.next();
    }
    String proxyHost = this.scanner.nextLine();
    if ((proxyHost == null) || ("".equals(proxyHost.trim())))
    {
      System.out.println("잘못 입력하셨습니다. Proxy Host를 입력하여 주십시요.");
      proxyHost = setProxyHost();
    }
    return proxyHost;
  }
  
  private int setProxyPort()
  {
    System.out.print("Proxy Port : ");
    while (!this.scanner.hasNextInt())
    {
      System.out.println("\n잘못 입력하셧습니다. Proxy Port를 입력하여 주십시요.");
      System.out.print("Proxy Port : ");
      this.scanner.nextLine();
    }
    return this.scanner.nextInt();
  }
  
  public static void main(String[] args) throws DependencyResolutionException, URISyntaxException, ModelBuildingException
  {	
	initialize();
  }
  
  private static void initialize() {
	  Scanner sc = null;
	    try
	    {
	      sc = new Scanner(System.in);
	      
	      System.out.println("========================================================================");
	      System.out.println("|     ___ _       _______    ______                      __            |");
	      System.out.println("|    /   | |     / / ___/   / ____/  ______  ____  _____/ /____  _____ |");
	      System.out.println("|   / /| | | /| / /\\__ \\   / __/ | |/_/ __ \\/ __ \\/ ___/ __/ _ \\/ ___/ |");
	      System.out.println("|  / ___ | |/ |/ /___/ /  / /____>  </ /_/ / /_/ / /  / /_/  __/ /     |");
	      System.out.println("| /_/  |_|__/|__//____/  /_____/_/|_/ .___/\\____/_/   \\__/\\___/_/      |");
	      System.out.println("|                                  /_/                                 |");
	      System.out.println("========================================================================");
	      System.out.println("");
	      
	      AWSExportStarter awsExportStarter = new AWSExportStarter(sc);
	      awsExportStarter.execute(); return;
	    }
	    catch (Exception e)
	    {
	      e.printStackTrace();
	    }
	    finally
	    {
	      if (sc != null) {
	        try
	        {
	          sc.close();
	        }
	        catch (Exception localException3) {}
	      }
	    }  
  }
  
  public Properties getProperties()
  {
    return this.properties;
  }
  
  public Scanner getScanner()
  {
    return this.scanner;
  }
}
