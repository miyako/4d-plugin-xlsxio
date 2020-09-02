/* --------------------------------------------------------------------------------
 #
 #  4DPlugin-XLSXIO.cpp
 #	source generated by 4D Plugin Wizard
 #	Project : XLSXIO
 #	author : miyako
 #	2020/09/02
 #  
 # --------------------------------------------------------------------------------*/

#include "4DPlugin-XLSXIO.h"

#pragma mark -

void PluginMain(PA_long32 selector, PA_PluginParameters params) {
    
	try
	{
        switch(selector)
        {
			// --- XLSXIO
            
			case 1 :
				XLSX_TO_JSON(params);
				break;
			case 2 :
				JSON_TO_XLSX(params);
				break;
			case 3:
				XLSX_SHEET_NAMES(params);
				break;
        }

	}
	catch(...)
	{

	}
}

#pragma mark -

void XLSX_TO_JSON(PA_PluginParameters params) {

    sLONG_PTR *pResult = (sLONG_PTR *)params->fResult;
    PackagePtr pParams = (PackagePtr)params->fParameters;

    C_TEXT Param1_path;
    C_LONGINT Param2_options;
    C_TEXT Param3_errors;
	C_TEXT Param3_sheet;
    C_TEXT returnValue;
    
    Param1_path.fromParamAtIndex(pParams, 1);
    Param2_options.fromParamAtIndex(pParams, 2);
	Param3_sheet.fromParamAtIndex(pParams, 3);

	CUTF8String _sheetname;
	Param3_sheet.copyUTF8String(&_sheetname);
	const char *sheetnamespecified = (const char *)_sheetname.c_str();
	bool issheetnamespecified = _sheetname.size();

    CUTF8String _path;
    Param1_path.copyPath(&_path);
    const char *path = (const char *)_path.c_str();
     
    unsigned int flags = Param2_options.getIntValue();
    
    Json::Value json(Json::objectValue);
    Json::Value json_sheets(Json::arrayValue);
    Json::Value json_errors(Json::objectValue);
    
    xlsxioreader xlsxioread = xlsxioread_open(path);
    
    if(!xlsxioread)
    {
        json_errors["error"] = "failed to open xlsx";
        json_errors["path"] = path;
        
    }else
    {
        xlsxioreadersheetlist sheetlist = xlsxioread_sheetlist_open(xlsxioread);
        
        if(!sheetlist)
        {
            json_errors["error"] = "failed to read sheet list";
            json_errors["path"] = path;
            
        }else
        {
            time_t startTime = time(0);
            
            const char *sheetname;
            while ((sheetname = xlsxioread_sheetlist_next(sheetlist)) != NULL)
            {
                Json::Value json_sheet(Json::objectValue);
                Json::Value json_sheet_rows(Json::arrayValue);
                
				xlsxioreadersheet sheet = NULL;

				if (issheetnamespecified)
				{
					sheet = xlsxioread_sheet_open(xlsxioread, sheetnamespecified, flags);
				}
				else
				{					
					sheet = xlsxioread_sheet_open(xlsxioread, sheetname, flags);
				}
                
                if(sheet)
                {
					json_sheet["name"] = sheetname;

                    while (xlsxioread_sheet_next_row(sheet))
                    {
						time_t now = time(0);
						time_t elapsedTime = abs(startTime - now);
						if (elapsedTime > 0)
						{
							startTime = now;
							PA_YieldAbsolute();
						}

                        Json::Value json_sheet_row(Json::objectValue);
                        Json::Value json_sheet_row_values(Json::arrayValue);
                        
                        char *value;
                        while ((value = xlsxioread_sheet_next_cell(sheet)) != NULL)
                        {
                            json_sheet_row_values.append(value);
                        }
                        
                        json_sheet_row["values"] = json_sheet_row_values;
                        json_sheet_rows.append(json_sheet_row);
                    }
                    
					json_sheet["rows"] = json_sheet_rows;
                    json_sheets.append(json_sheet);
                    
                    xlsxioread_sheet_close(sheet);

					if (issheetnamespecified) break;

                }else
                {
                    json_errors["error"] = "failed to read sheet";
                    json_errors["path"] = path;
                    break;
                }
                
            }
			json["sheets"] = json_sheets;
            xlsxioread_sheetlist_close(sheetlist);
        }
        xlsxioread_close(xlsxioread);
    }
    
    Json::StreamWriterBuilder writer;
	writer["indentation"] = "";

    std::string errors = Json::writeString(writer, json_errors);
    Param3_errors.setUTF8String((const uint8_t *)errors.c_str(), (uint32_t)errors.length());
    Param3_errors.toParamAtIndex(pParams, 3);
    
    std::string status = Json::writeString(writer, json);
    returnValue.setUTF8String((const uint8_t *)status.c_str(), (uint32_t)status.length());
    returnValue.setReturn(pResult);
}

void JSON_TO_XLSX(PA_PluginParameters params) {

    sLONG_PTR *pResult = (sLONG_PTR *)params->fResult;
    PackagePtr pParams = (PackagePtr)params->fParameters;
    
    C_TEXT Param1_path;
    C_TEXT Param2_json;
    C_LONGINT Param3_row_height;
    C_LONGINT Param4_detection_rows;
    C_TEXT Param5_errors;
    
    Param1_path.fromParamAtIndex(pParams, 1);
    Param2_json.fromParamAtIndex(pParams, 2);
    Param3_row_height.fromParamAtIndex(pParams, 3);
    Param4_detection_rows.fromParamAtIndex(pParams, 4);
    
    CUTF8String _path;
    Param1_path.copyPath(&_path);
    const char *path = (const char *)_path.c_str();
    
    CUTF8String _json;
	Param2_json.copyUTF8String(&_json);
    
    Json::Value json;
    Json::Value json_errors;
    
    Json::CharReaderBuilder builder;
    std::string errors;
    
    Json::CharReader *reader = builder.newCharReader();
    bool parse = reader->parse((const char *)_json.c_str(),
        (const char *)_json.c_str() + _json.size(),
        &json,
        &errors);
    delete reader;
    
    if(!parse){
        
        json_errors["error"] = "failed to parse json";
        json_errors["path"] = path;
        
    }else{
        
        Json::Value json_sheets = json["sheets"];
        
        if(json_sheets)
        {
           
            if(json_sheets.isArray())
            {
                if(json_sheets.size())
                {
                    Json::Value json_sheet = json_sheets[0];
                    
                    if((json_sheet) && json_sheet.isObject())
                    {
                        Json::Value n = json_sheet["name"];
                        if(n)
                        {
                            if(n.isString())
                            {
                                std::string sheetname = n.asString();
                                
                                xlsxiowriter xlsxiowrite = xlsxiowrite_open(
                                                                            (const char *)path,
                                                                            (const char *)sheetname.c_str());
                                if (!xlsxiowrite)
                                {
                                    json_errors["error"] = "failed to create xlsx";
                                    json_errors["path"] = path;
                                }else
                                {
                                    xlsxiowrite_set_row_height(xlsxiowrite,
                                                               Param3_row_height.getIntValue());
                                    xlsxiowrite_set_detection_rows(xlsxiowrite,
                                                                   Param4_detection_rows.getIntValue());
                                    
                                    Json::Value json_sheet_rows = json_sheet["rows"];
                                    
                                    if(json_sheet_rows)
                                    {
                                        if(json_sheet_rows.isArray())
                                        {
											time_t startTime = time(0);

                                            for(unsigned int i = 0; i < json_sheet_rows.size();++i)
                                            {
												time_t now = time(0);
												time_t elapsedTime = abs(startTime - now);
												if (elapsedTime > 0)
												{
													startTime = now;
													PA_YieldAbsolute();
												}

                                                Json::Value json_sheet_row = json_sheet_rows[i];
                                                if(json_sheet_row.isObject())
                                                {
                                                    Json::Value json_sheet_row_values = json_sheet_row["values"];
                                                    
                                                    if(json_sheet_row_values)
                                                    {
                                                        if(json_sheet_row_values.isArray())
                                                        {
                                                            for(unsigned int j = 0; j < json_sheet_row_values.size();++j)
                                                            {
                                                                Json::Value json_sheet_row_value = json_sheet_row_values[j];
                                                                
                                                                Json::ValueType type = json_sheet_row_value.type();
                                                                
                                                                switch (type)
                                                                {
                                                                    case Json::nullValue:
                                                                        xlsxiowrite_add_cell_string(xlsxiowrite, NULL);
                                                                        break;
                                                                    case Json::stringValue:
                                                                    {
                                                                        std::string stringValue = json_sheet_row_value.asString();
                                                                        xlsxiowrite_add_cell_string(xlsxiowrite, (const char *)stringValue.c_str());
                                                                    }
                                                                        break;
                                                                    case Json::intValue:
                                                                        xlsxiowrite_add_cell_int(xlsxiowrite, json_sheet_row_value.asInt());
                                                                        break;
                                                                    case Json::uintValue:
                                                                            xlsxiowrite_add_cell_int(xlsxiowrite, json_sheet_row_value.asUInt());
                                                                            break;
                                                                    case Json::realValue:
                                                                        xlsxiowrite_add_cell_float(xlsxiowrite, json_sheet_row_value.asDouble());
                                                                        break;
                                                                    case Json::booleanValue:
                                                                        xlsxiowrite_add_cell_int(xlsxiowrite, json_sheet_row_value.asBool());
                                                                        break;
                                                                    case Json::arrayValue:
                                                                    case Json::objectValue:
                                                                    default:
                                                                        json_errors["error"] = "sheets[0].rows[].values[] must be null, string, number or bool";
                                                                        json_errors["row"] = i;
                                                                        json_errors["col"] = j;
                                                                        break;
                                                                }
                                                                
                                                            }
                                                        }else
                                                        {
                                                            json_errors["error"] = "sheets[0].rows[].values is not an array";
                                                            json_errors["row"] = i;
                                                        }
                                                    }else
                                                    {
                                                        json_errors["error"] = "sheets[0].rows[].values is missing";
                                                        json_errors["row"] = i;
                                                    }
                                                }
												xlsxiowrite_next_row(xlsxiowrite);
                                            }
                                        }else
                                        {
                                            json_errors["error"] = "sheets[0].rows is not an array";
                                        }
                                    }else{
                                        json_errors["error"] = "sheets[0].rows is missing";
                                    }
									
									xlsxiowrite_close(xlsxiowrite);
                                }
  
                            }else
                            {
                                json_errors["error"] = "sheets[0].name must be a string";
                            }
 
                        }else
                        {
                            json_errors["error"] = "sheets[0].name is missing";
                        }
                        
                    }
                    
                }else
                {
                    json_errors["error"] = "sheets[0] is missing";
                }
            }else
            {
                json_errors["error"] = "sheets is not an array";
            }
            
        }else
        {
            json_errors["error"] = "sheets[0] is missing";
        }
        
    }

    Json::StreamWriterBuilder writer;
	writer["indentation"] = "";

    errors = Json::writeString(writer, json_errors);
    Param5_errors.setUTF8String((const uint8_t *)errors.c_str(), (uint32_t)errors.length());
    Param5_errors.toParamAtIndex(pParams, 5);
}

void XLSX_SHEET_NAMES(PA_PluginParameters params)
{
	PackagePtr pParams = (PackagePtr)params->fParameters;

	C_TEXT Param1_path;

	Param1_path.fromParamAtIndex(pParams, 1);

	CUTF8String _path;
	Param1_path.copyPath(&_path);
	const char *path = (const char *)_path.c_str();

	xlsxioreader xlsxioread = xlsxioread_open(path);

	PA_Variable names = PA_CreateVariable(eVK_ArrayUnicode);
	PA_ResizeArray(&names, 0);

	if (xlsxioread)
	{
		xlsxioreadersheetlist sheetlist = xlsxioread_sheetlist_open(xlsxioread);

		if (sheetlist)
		{
			const char *sheetname;
			while ((sheetname = xlsxioread_sheetlist_next(sheetlist)) != NULL)
			{
				C_TEXT t;
				t.setUTF8String((const uint8_t *)sheetname, strlen(sheetname));

				PA_Unistring ustr = PA_CreateUnistring((PA_Unichar *)t.getUTF16StringPtr());
				PA_long32 size = PA_GetArrayNbElements(names);
				size++;
				PA_ResizeArray(&names, size);
				PA_SetStringInArray(names, size, &ustr);
			}

			xlsxioread_sheetlist_close(sheetlist);
		}

		xlsxioread_close(xlsxioread);
	}

	PA_VariableKind paramKind = PA_GetVariableKind(PA_GetVariableParameter(params, 2));
		
	if ((paramKind == eVK_ArrayUnicode) || (paramKind == eVK_Undefined))
	{
		PA_SetVariableParameter(params, 2, names, 0);
	}
	
}