// JavaScript Document

function InitArea(ObjProvince , ObjCity , ObjCounty , ObjProvinceArray , ObjCityArray , ObjCountyArray , strProvince , strCity , strCounty)
{
	var thisProvince = document.getElementById(ObjProvince);
	var thisCity = document.getElementById(ObjCity);
	var thisCounty = document.getElementById(ObjCounty);
	
	thisProvince.options[0] = new Option("省份" , "");
	thisCity.options[0] = new Option("地级市" , "");
	thisCounty.options[0] = new Option("不限" , "不限");
	
	for(var i = 0; i < ObjProvinceArray.length; i++)
	{
		thisProvince.options[i+1] = new Option(ObjProvinceArray[i] , ObjProvinceArray[i]);
		if(ObjProvinceArray[i] == strProvince)
			thisProvince.options[i+1].selected = true;
	}
	
	if(strProvince != "")
	{
		for(i = 0; i < ObjProvinceArray.length; i++)
		{
			if(ObjProvinceArray[i] == strProvince)
			{
				for(var j = 0; j < ObjCityArray[i].length; j++)
				{
					thisCity.options[j + 1] = new Option(ObjCityArray[i][j], ObjCityArray[i][j]);
					if(ObjCityArray[i][j] == strCity)
						thisCity.options[j + 1].selected = true;
				}
			}
		}

	}
	
	if(strCity != "")
	{
		for(i = 0; i < ObjCityArray.length; i++)
		{
			for(j = 0; j < ObjCityArray[i].length; j++)
			{
				if(ObjCityArray[i][j] == strCity)
				{
					for(var k = 0; k < ObjCountyArray[i][j].length; k++)
					{
						thisCounty.options[k + 1] = new Option(ObjCountyArray[i][j][k], ObjCountyArray[i][j][k]);
						if(ObjCountyArray[i][j][k] == strCounty)
							thisCounty.options[k + 1].selected = true;
					}
				}
			}
		}

	}
	
}

function FzInitArea(ObjProvince , ObjCity , ObjCounty , ObjProvinceArray , ObjCityArray , ObjCountyArray , strProvince , strCity , strCounty)
{
	var thisProvince = document.getElementById(ObjProvince);
	var thisCity = document.getElementById(ObjCity);
	var thisCounty = document.getElementById(ObjCounty);
	
	//thisProvince.options[0] = new Option("省份" , "");
	//thisCity.options[0] = new Option("地级市" , "");
	//thisCounty.options[0] = new Option("不限" , "不限");
	
	for(var i = 0; i < ObjProvinceArray.length; i++)
	{
		//thisProvince.options[i+1] = new Option(ObjProvinceArray[i] , ObjProvinceArray[i]);
		if(ObjProvinceArray[i] == strProvince)
		{
			thisProvince.options[0] = new Option(ObjProvinceArray[i] , ObjProvinceArray[i]);
			thisProvince.options[0].selected = true;
		}
	}
	
	if(strProvince != "")
	{
		for(i = 0; i < ObjProvinceArray.length; i++)
		{
			if(ObjProvinceArray[i] == strProvince)
			{
				if(strCity == "")
				{
					thisCity.options[0] = new Option("地级市" , "");
				}
				for(var j = 0; j < ObjCityArray[i].length; j++)
				{
					if(strCity == "")
					{
						thisCity.options[j + 1] = new Option(ObjCityArray[i][j], ObjCityArray[i][j]);
						if(ObjCityArray[i][j] == strCity)
						{
							thisCity.options[j + 1].selected = true;
						}
					}
					else
					{
						if(ObjCityArray[i][j] == strCity)
						{
							thisCity.options[0] = new Option(ObjCityArray[i][j], ObjCityArray[i][j]);
							thisCity.options[0].selected = true;
						}
					}
				}
			}
		}

	}
	
	if(strCity != "")
	{
		for(i = 0; i < ObjCityArray.length; i++)
		{
			for(j = 0; j < ObjCityArray[i].length; j++)
			{
				if(ObjCityArray[i][j] == strCity)
				{
					if(strCounty == "")
					{
						thisCounty.options[0] = new Option("不限" , "不限");
					}
					for(var k = 0; k < ObjCountyArray[i][j].length; k++)
					{
						if(strCounty == "" || strCounty == "不限")
						{
							thisCounty.options[k + 1] = new Option(ObjCountyArray[i][j][k], ObjCountyArray[i][j][k]);
							if(ObjCountyArray[i][j][k] == strCounty)
							{
								thisCounty.options[k + 1].selected = true;
							}
						}
						else
						{
							if(ObjCountyArray[i][j][k] == strCounty)
							{
								thisCounty.options[0] = new Option(ObjCountyArray[i][j][k], ObjCountyArray[i][j][k]);
								thisCounty.options[0].selected = true;
							}
						}
					}
				}
			}
		}
	}
}

function SelChgCity(ObjCityField , ProvinceValue , ObjProvinceArray , ObjCityArray , ObjCountyField)
{
	var thisCityField = document.getElementById(ObjCityField);
	var thisCountyField = document.getElementById(ObjCountyField);
	for(var i = thisCityField.length; i > 0; i--)
	{
		thisCityField.remove(i);
	}
	for(var i = thisCountyField.length; i > 0; i--)
	{
		thisCountyField.remove(i);
	}
	thisCityField.options[0] = new Option("地级市" , "");
	thisCountyField.options[0] = new Option("不限" , "不限");
	if(ProvinceValue != "")
	{
		for(i = 0; i < ObjProvinceArray.length; i++)
		{
			if(ObjProvinceArray[i] == ProvinceValue)
			{
				for(var j = 0; j < ObjCityArray[i].length; j++)
				{
					thisCityField.options[j + 1] = new Option(ObjCityArray[i][j], ObjCityArray[i][j]);
				}
			}
		}

	}
}

function SelChgCounty(ObjCountyField , CityValue , ObjCityArray , ObjCountyArray)
{
	var thisCountyField = document.getElementById(ObjCountyField);
	for(var i = thisCountyField.length; i > 0; i--)
	{
		thisCountyField.remove(i);
	}
	thisCountyField.options[0] = new Option("不限" , "不限");
	if(CityValue != "")
	{
		for(i = 0; i < ObjCityArray.length; i++)
		{
			for(var j = 0; j < ObjCityArray[i].length; j++)
			{
				if(ObjCityArray[i][j] == CityValue)
				{
					for(var k = 0; k < ObjCountyArray[i][j].length; k++)
					{
						thisCountyField.options[k + 1] = new Option(ObjCountyArray[i][j][k], ObjCountyArray[i][j][k]);
					}
				}
			}
		}

	}
}