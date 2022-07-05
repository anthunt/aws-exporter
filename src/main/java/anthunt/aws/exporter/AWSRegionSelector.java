package anthunt.aws.exporter;

import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Set;

import software.amazon.awssdk.regions.Region;

public class AWSRegionSelector
{
  private static final String FAV_PROP_PREFIX = "favorite.";
  private HashMap<Integer, List<Region>> favRegions;
  private AWSExportStarter awsExportStarter;
  private List<Region> allRegions;
  private List<Region> selectedRegions;
  
  public AWSRegionSelector(AWSExportStarter awsExportStarter)
  {
	this.favRegions = new HashMap<>();
    this.selectedRegions = new ArrayList<>();
    this.allRegions = new ArrayList<>();
    this.awsExportStarter = awsExportStarter;
    
    List<Region> allRegions = Region.regions();
    for (Region region : allRegions) {
      this.allRegions.add(region);
    }
    select();
  }
  
  public static String padRight(String str, int length, String padChar)
  {
    String pad = "";
    for (int i = 0; i < length; i++) {
      pad = pad + padChar;
    }
    return str + pad.substring(str.length());
  }
  
  private void findFavoriteRegions()
  {
    Set<Object> keys = this.awsExportStarter.getProperties().keySet();
    List<String> favKeys = new ArrayList<>();
    for (Object keyObject : keys)
    {
      String key = (String)keyObject;
      if (key.startsWith(FAV_PROP_PREFIX)) {
        favKeys.add(key);
      }
    }

    int iFav;
    if (favKeys.size() > 0)
    {
      System.out.println("-------------------------------------------------");
      System.out.println("Select Favorite Region");
      System.out.println("");
      
      iFav = 100;
      for (String key : favKeys)
      {
        System.out.print("[" + iFav + "] " + key.replaceAll(FAV_PROP_PREFIX, "") + " : ");
        String propValue = this.awsExportStarter.getProperties().getProperty(key);
        StringBuffer valueBuffer = new StringBuffer();
        String[] values = propValue.split(",");
        
        List<Region> fav = new ArrayList<>();
        for (String value : values)
        {
          value = value.trim();
          if (valueBuffer.length() > 0) {
            valueBuffer.append(", ");
          }
          valueBuffer.append(value);
          fav.add(Region.of(value));
        }
        this.favRegions.put(Integer.valueOf(iFav), fav);
        System.out.println(valueBuffer.toString());
        iFav++;
      }
    }
  }
  
  private void showRegions()
  {
    for (int i = 0; i < this.allRegions.size(); i++)
    {
      Region region = (Region)this.allRegions.get(i);
      if ((i + 1) % 2 == 0) {
        System.out.println("[" + (i + 1) + "] " + region.id());
      } else {
        System.out.print(padRight("[" + (i + 1) + "] " + region.id(), 30, " "));
      }
    }
    System.out.println("");
  }
  
  private void select()
  {
    System.out.println("");
    System.out.println("=================================================");
    System.out.println("Select an extraction region. Multiple selection possible [Comma(,)]");
    System.out.println("-------------------------------------------------");
    System.out.println("[0] Create Favorite Extract Regions");
    System.out.println("-------------------------------------------------");
    System.out.println("Select individual regions");
    System.out.println("");
    showRegions();
    findFavoriteRegions();
    System.out.println("-------------------------------------------------");
    System.out.println("[99] All Regions");
    System.out.println("=================================================");
    
    selectInput();
    
    System.out.println("-------------------------------------------------");
    System.out.println("Selected region [" + this.selectedRegions.size() + "ea]");
    System.out.println("-------------------------------------------------");
    for (Region region : this.selectedRegions) {
      System.out.println(region.id());
    }
    System.out.println("-------------------------------------------------");
  }
  
  private void makeFavRegions()
  {
    System.out.println("Create Favorite Region Collection");
    String favName = setFavRegionsName();
    setFavRegions(favName);
  }
  
  private void setFavRegions(String favName)
  {
    System.out.println("=================================================");
    System.out.println("Please select your favorite region.");
    System.out.println("-------------------------------------------------");
    showRegions();
    System.out.println("=================================================");
    
    StringBuffer favRegions = selectFavInput(favName);
    this.awsExportStarter.getProperties().put(FAV_PROP_PREFIX + favName, favRegions.toString());
    this.awsExportStarter.storeProperties();
    
    select();
  }
  
  private StringBuffer selectFavInput(String favName)
  {
    StringBuffer favBuffer = new StringBuffer();
    System.out.println("");
    System.out.print("Select Favorite Region : ");
    String input = this.awsExportStarter.getScanner().nextLine();
    System.out.println("");
    if (input.length() > 0)
    {
      int exceptionCount = 0;
      String[] selected = input.split(",");
      for (String selectedIndex : selected) {
        try
        {
          int selectedIdx = Integer.parseInt(selectedIndex.trim());
          if ((selectedIdx < 1) || (selectedIdx > this.allRegions.size())) {
            throw new RegionSelectorException("[" + selectedIdx + "] is an optional value that is not supported..");
          }
          if (favBuffer.length() > 0) {
            favBuffer.append(",");
          }
          favBuffer.append(((Region)this.allRegions.get(selectedIdx - 1)).id());
        }
        catch (NumberFormatException e)
        {
          exceptionCount++;
          System.out.println("[" + selectedIndex + "] is an optional value that is not supported.");
        }
        catch (Exception e)
        {
          exceptionCount++;
          System.out.println(e.getMessage());
        }
      }
      if (exceptionCount > 0) {
        favBuffer = selectFavInput(favName);
      }
    }
    else
    {
      favBuffer = selectFavInput(favName);
    }
    return favBuffer;
  }
  
  private String setFavRegionsName()
  {
    System.out.print("Favorite Name(English) : ");
    String favName = this.awsExportStarter.getScanner().nextLine();
    System.out.println("");
    if ((favName == null) || ("".equals(favName.trim())))
    {
      System.out.println("Please enter your favorite name.");
      favName = setFavRegionsName();
    }
    return favName;
  }
  
  private void selectInput()
  {
    this.selectedRegions.clear();
    System.out.println("");
    System.out.print("Extract Region Selection : ");
    String input = this.awsExportStarter.getScanner().nextLine();
    System.out.println("");
    if (input.length() > 0)
    {
      int exceptionCount = 0;
      String[] selected = input.split(",");
      for (String selectedIndex : selected) {
        try
        {
          int selectedIdx = Integer.parseInt(selectedIndex.trim());
          if (selectedIdx == 0)
          {
            makeFavRegions();
            break;
          }
          if (selectedIdx == 99)
          {
            this.selectedRegions.clear();
            this.selectedRegions.addAll(this.allRegions);
            break;
          }
          if ((selectedIdx >= 100) && (selectedIdx < 100 + this.favRegions.size()))
          {
            this.selectedRegions.clear();
            this.selectedRegions.addAll(this.favRegions.get(Integer.valueOf(selectedIdx)));
            break;
          }
          if ((selectedIdx < 1) || (selectedIdx > this.allRegions.size())) {
            throw new RegionSelectorException("[" + selectedIdx + "] is an optional value that is not supported.");
          }
          this.selectedRegions.add(this.allRegions.get(selectedIdx - 1));
        }
        catch (NumberFormatException e)
        {
          exceptionCount++;
          System.out.println("[" + selectedIndex + "] is an optional value that is not supported.");
        }
        catch (Exception e)
        {
          exceptionCount++;
          System.out.println(e.getMessage());
        }
      }
      if (exceptionCount > 0) {
        selectInput();
      }
    }
    else
    {
      selectInput();
    }
  }
  
  public List<Region> getRegions()
  {
    return this.selectedRegions;
  }
  
  public void clear()
  {
    this.allRegions.clear();
    this.selectedRegions.clear();
  }
}
