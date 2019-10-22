package anthunt.aws.exporter;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.net.URISyntaxException;
import java.util.ArrayList;
import java.util.Enumeration;
import java.util.Iterator;
import java.util.List;
import java.util.Map;
import java.util.Properties;
import java.util.Scanner;
import java.util.Set;

import org.apache.maven.model.building.ModelBuildingException;
import org.eclipse.aether.resolution.DependencyResolutionException;

import anthunt.aws.exporter.model.AmazonAccess;
import software.amazon.awssdk.profiles.Profile;
import software.amazon.awssdk.profiles.ProfileFile;
import software.amazon.awssdk.regions.Region;

public class AWSExportStarter
{
  private static final String dataFileName = "conf/awsExporter.data";
  private Scanner scanner;
  private Properties properties;
  private File propertiesFile;
  private AmazonAccess amazonAccess;
  
  private static enum PropType
  {
       USE_PROXY("connect.useProxy")
    ,  PROXY_HOST("connect.proxy.host")
    ,  PROXY_PORT("connect.proxy.port")
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
      AmazonAccess amazonAccess = new AmazonAccess();
      
      switch (propType)
      {
      case USE_PROXY: 
        amazonAccess.setUseProxy(Boolean.valueOf(Boolean.parseBoolean(value)));
        break;
      case PROXY_HOST: 
        amazonAccess.setProxyHost(value);
        break;
      case PROXY_PORT: 
        amazonAccess.setProxyPort(Integer.valueOf(Integer.parseInt(value)));
        break;
      }
      
      this.amazonAccess = amazonAccess;
      
    }
  }
  
  private void showAccessList()
  {

      List<String> keyIndex = new ArrayList<>();
      
      System.out.println("Select your profile.\n");
      System.out.println("==============================");
      
      ProfileFile profileFile = ProfileFile.defaultProfileFile();
      
      Map<String, Profile> profileMap = profileFile.profiles();
      
      Set<String> keySet = profileMap.keySet();
      Iterator<String> keys = keySet.iterator();
      while(keys.hasNext()) {
    	  String profileName = keys.next();
    	  keyIndex.add(profileName);
    	  System.out.print(keyIndex.size() + ". " + profileName);
          System.out.println("");
      }
      System.out.println("------------------------------");
      System.out.println("99. Exit");
      System.out.println("==============================");
      System.out.println("");
      
      checkSelectAccessList(keyIndex);
    
  }
  
  private void exit() {
	  System.exit(-1);
  }
  
  private void checkSelectAccessList(List<String> keyIndex)
  {
    System.out.print("Profile Number : ");
    while (!this.scanner.hasNextInt())
    {
      System.out.println("\nWrong number. Try again.");
      System.out.print("Profile Number : ");
      this.scanner.nextLine();
    }
    int index = this.scanner.nextInt();
    this.scanner.nextLine();
    if (index == 99)
    {
      this.exit();
    }
    else if ((index < 1) || (index > keyIndex.size()))
    {
      System.out.println("Wrong number. Try again.");
      checkSelectAccessList(keyIndex);
    }
    else
    {
      this.extractRegion(keyIndex.get(index - 1));      
    }
  }
  
  private void extractRegion(String profileName) {
	  List<Region> regions = new AWSRegionSelector(this).getRegions();
      
      System.out.println("");
      System.out.println(profileName + " AWS 정보 생성 시작");
      new AWSExporter(this.amazonAccess, regions, profileName);
      System.out.println(profileName + " AWS 정보 생성 완료");
      System.out.println("");
      
      showAccessList();
  }
  
  private void makeAccess()
  {
    System.out.println("\n접속 정보를 입력하여 주십시요.");
        
    AmazonAccess amazonAccess = new AmazonAccess();
        
    boolean isUseProxy = setUseProxy();
    amazonAccess.setUseProxy(Boolean.valueOf(isUseProxy));
    if (isUseProxy)
    {
      String proxyHost = setProxyHost();
      int proxyPort = setProxyPort();
      amazonAccess.setProxyHost(proxyHost);
      amazonAccess.setProxyPort(Integer.valueOf(proxyPort));
    }
    
    this.amazonAccess = amazonAccess;
    
    this.properties.put(PropType.USE_PROXY.getPostFix(), amazonAccess.isUseProxy().toString());
    
    if(amazonAccess.isUseProxy()) {
	    this.properties.put(PropType.PROXY_HOST.getPostFix(), amazonAccess.getProxyHost());
	    this.properties.put(PropType.PROXY_PORT.getPostFix(), amazonAccess.getProxyPort().toString());
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
