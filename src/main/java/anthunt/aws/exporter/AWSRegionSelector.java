package anthunt.aws.exporter;

import com.amazonaws.regions.Regions;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Set;

public class AWSRegionSelector
{
  private static final String FAV_PROP_PREFIX = "favorite.";
  private HashMap<Integer, List<Regions>> favRegions;
  private AWSExportStarter awsExportStarter;
  private List<Regions> allRegions;
  private List<Regions> selectedRegions;
  
  public AWSRegionSelector(AWSExportStarter awsExportStarter)
  {
	this.favRegions = new HashMap<>();
    this.selectedRegions = new ArrayList<>();
    this.allRegions = new ArrayList<>();
    this.awsExportStarter = awsExportStarter;
    
    Regions[] allRegions = Regions.values();
    for (Regions regions : allRegions) {
      this.allRegions.add(regions);
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
      System.out.println("즐겨찾기 Region 선택");
      System.out.println("");
      
      iFav = 100;
      for (String key : favKeys)
      {
        System.out.print("[" + iFav + "] " + key.replaceAll(FAV_PROP_PREFIX, "") + " : ");
        String propValue = this.awsExportStarter.getProperties().getProperty(key);
        StringBuffer valueBuffer = new StringBuffer();
        String[] values = propValue.split(",");
        
        List<Regions> fav = new ArrayList<>();
        for (String value : values)
        {
          value = value.trim();
          if (valueBuffer.length() > 0) {
            valueBuffer.append(", ");
          }
          valueBuffer.append(value);
          fav.add(Regions.fromName(value));
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
      Regions regions = (Regions)this.allRegions.get(i);
      if ((i + 1) % 2 == 0) {
        System.out.println("[" + (i + 1) + "] " + regions.getName());
      } else {
        System.out.print(padRight("[" + (i + 1) + "] " + regions.getName(), 30, " "));
      }
    }
  }
  
  private void select()
  {
    System.out.println("");
    System.out.println("=================================================");
    System.out.println("추출 Region을 선택하여 주십시요. 다중선택 가능 [콤마(,) 구분]");
    System.out.println("-------------------------------------------------");
    System.out.println("[0] 즐겨찾기 추출 Region 생성");
    System.out.println("-------------------------------------------------");
    System.out.println("개별 Region 선택");
    System.out.println("");
    showRegions();
    findFavoriteRegions();
    System.out.println("-------------------------------------------------");
    System.out.println("[99] All Regions");
    System.out.println("=================================================");
    
    selectInput();
    
    System.out.println("-------------------------------------------------");
    System.out.println("선택된 Region [" + this.selectedRegions.size() + "개]");
    System.out.println("-------------------------------------------------");
    for (Regions regions : this.selectedRegions) {
      System.out.println(regions.getName());
    }
    System.out.println("-------------------------------------------------");
  }
  
  private void makeFavRegions()
  {
    System.out.println("즐겨찾기 Region 모음 생성");
    String favName = setFavRegionsName();
    setFavRegions(favName);
  }
  
  private void setFavRegions(String favName)
  {
    System.out.println("=================================================");
    System.out.println("즐겨찾기 Region을 선택하여 주십시요.");
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
    System.out.print("즐겨찾기 Region 선택 : ");
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
            throw new Exception("[" + selectedIdx + "] 는 지원하지 않는 선택 값 입니다..");
          }
          if (favBuffer.length() > 0) {
            favBuffer.append(",");
          }
          favBuffer.append(((Regions)this.allRegions.get(selectedIdx - 1)).getName());
        }
        catch (NumberFormatException e)
        {
          exceptionCount++;
          System.out.println("[" + selectedIndex + "] 는 지원하지 않는 선택 값 입니다.");
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
    System.out.print("즐겨찾기 명(영문) : ");
    String favName = this.awsExportStarter.getScanner().nextLine();
    System.out.println("");
    if ((favName == null) || ("".equals(favName.trim())))
    {
      System.out.println("즐겨찾기 명을 입력하여 주십시요.");
      favName = setFavRegionsName();
    }
    return favName;
  }
  
  private void selectInput()
  {
    this.selectedRegions.clear();
    System.out.println("");
    System.out.print("추출 Region 선택 : ");
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
            throw new Exception("[" + selectedIdx + "] 는 지원하지 않는 선택 값 입니다.");
          }
          this.selectedRegions.add(this.allRegions.get(selectedIdx - 1));
        }
        catch (NumberFormatException e)
        {
          exceptionCount++;
          System.out.println("[" + selectedIndex + "] 는 지원하지 않는 선택 값 입니다.");
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
  
  public List<Regions> getRegions()
  {
    return this.selectedRegions;
  }
  
  public void clear()
  {
    this.allRegions.clear();
    this.selectedRegions.clear();
  }
}
