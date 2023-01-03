package com.focus.supporters;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.LinkedHashMap;
import java.util.Map;
import java.util.Properties;
import java.util.Set;

import com.focus.base.BaseEngine;

public class PropertiesReusables extends BaseEngine
{

	private String filePath;
	private FileInputStream fip;
	private Properties properties;
	
	public PropertiesReusables(String filePath) throws IOException
	{
		fip=new FileInputStream(filePath);
		properties=new Properties();
		properties.load(fip);
		
		if(properties!=null)
		{
			getLogger().info("Properties Object is available");
		}
		else
		{
			getLogger().info("Properties Object is Null, so It is not available");
		}
		
	}
	
	
	
	public String getPropertyValue(String key)
	{
		String value=properties.getProperty(key);
		return value;
	}
	
	
	
	public String getPropertyValue(String key, String defaultValue)
	{
		String value=properties.getProperty(key);
		if(value==null)
		{
			value=defaultValue;
		    getLogger().info("key is not available with value : " +value);
		}
		else
		{
		    getLogger().info("key is available with value : " +value);
		}
		return value;
	
	}
	
	
	public String getProperty(Object key)
	{
		String value=(String)properties.get(key);
		return value;
	}
	
	
	public String getProperty(Object key,Object defaultValue)
	{
		String value=(String)properties.get(key);
		return value;
	}
	
	
	public void setPropertyKeyValue(String key, String value) throws IOException
	{
		properties.setProperty(key, value);
		FileOutputStream fop=new FileOutputStream(filePath);
		properties.store(fop,"Data added Successully");
		
	}
	
	
	public void removeEntry(Object key) throws IOException
	{
		properties.remove(key);
		FileOutputStream fop=new FileOutputStream(filePath);
		properties.store(fop,"Remove file and Save Successully");
	}
	
	
	public void removeEntry(Object key, Object value) throws IOException
	{
		properties.remove(key, value);
		FileOutputStream fop=new FileOutputStream(filePath);
		properties.store(fop,"Remove file and Save Successfully");
		
	}
	
	public Map<String, String> getAllEntries()
	{
		Map<String,String> map=new LinkedHashMap<>();
		
		Set<Object> keys=properties.keySet();
		for(Object key : keys)
		{
			String key1=(String) key;
			String value=properties.getProperty(key1);
			map.put(key1, value);
		}
		
		return map;
	}
	
	
	
	
	public static void main(String[] args) throws IOException
	{
		PropertiesReusables pr=new PropertiesReusables("");
		Map<String, String> entries=pr.getAllEntries();
		
		Set<String> keys=entries.keySet();
		
		for(String key:keys)
		{
			if(key.equalsIgnoreCase("job")||key.equalsIgnoreCase("practice"))
			{
				String value=pr.getProperty(key);						
			}
		}
			
	}
	
	
}
