package com.thingworx.extension.custom.msoffice;

import com.thingworx.data.util.InfoTableInstanceFactory;
import com.thingworx.entities.utils.ThingUtilities;
import com.thingworx.logging.LogUtilities;
import com.thingworx.metadata.FieldDefinition;
import com.thingworx.metadata.annotations.ThingworxServiceDefinition;
import com.thingworx.metadata.annotations.ThingworxServiceParameter;
import com.thingworx.metadata.annotations.ThingworxServiceResult;
import com.thingworx.resources.Resource;
import com.thingworx.things.repository.FileRepositoryThing;
import com.thingworx.types.InfoTable;
import com.thingworx.types.collections.ValueCollection;
import com.thingworx.types.primitives.BooleanPrimitive;
import com.thingworx.types.primitives.NumberPrimitive;
import com.thingworx.types.primitives.StringPrimitive;
import java.io.FileInputStream;
import java.io.IOException;
import java.util.stream.Collectors;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.slf4j.Logger;

public class MSOfficeResource extends Resource {

  private final static Logger SCRIPT_LOGGER = LogUtilities.getInstance().getScriptLogger(MSOfficeResource.class);
  private static final long serialVersionUID = 1L;

  @ThingworxServiceDefinition(name = "readExcel", description = "", category = "", isAllowOverride = false, aspects = {"isAsync:false"})
  @ThingworxServiceResult(name = "response", description = "", baseType = "INFOTABLE", aspects = {"isEntityDataShape:true"})
  public InfoTable readExcel(
          @ThingworxServiceParameter(name = "fileRepository", description = "", baseType = "THINGNAME", aspects = {"isRequired:true", "thingTemplate:FileRepository"}) String fileRepository,
          @ThingworxServiceParameter(name = "path", description = "", baseType = "STRING", aspects = {"isRequired:true"}) String path,
          @ThingworxServiceParameter(name = "sheetIndex", description = "the index of the sheet, default = 0", baseType = "INTEGER") Integer sheetIndex,
          @ThingworxServiceParameter(name = "hasHeader", description = "true if the sheet has an header, false otherwise, default = false", baseType = "BOOLEAN") Boolean hasHeader,
          @ThingworxServiceParameter(name = "dataShape", description = "the output DataShape", baseType = "DATASHAPENAME", aspects = {"isRequired:true"}) String dataShape) throws Exception {
    SCRIPT_LOGGER.debug("MSOfficeResource - readExcel -> Start");

    InfoTable response = InfoTableInstanceFactory.createInfoTableFromDataShape(dataShape);
    @SuppressWarnings("CollectionsToArray")
    String[] names = response.getDataShape().getFields().getOrderedFieldsByOrdinal().stream().map(FieldDefinition::getName).collect(Collectors.toList()).toArray(new String[0]);

    FileRepositoryThing FileRepositoryThing = (FileRepositoryThing) ThingUtilities.findThing(fileRepository);
    try (FileInputStream fis = FileRepositoryThing.openFileForRead(path)) {
      Workbook wb;
      try {
        wb = new XSSFWorkbook(fis);
      } catch (IOException ex) {
        wb = new HSSFWorkbook(fis);
      }

      final boolean header = hasHeader != null ? hasHeader : false;
      wb.getSheetAt(sheetIndex != null ? sheetIndex : 0).forEach(row -> {
        if (!header || row.getRowNum() != 0) {
          ValueCollection values = new ValueCollection();

          row.forEach(cell -> {
            switch (cell.getCellTypeEnum()) {
              case STRING:
                values.put(names[cell.getColumnIndex()], new StringPrimitive(cell.getStringCellValue()));
                break;
              case NUMERIC:
                values.put(names[cell.getColumnIndex()], new NumberPrimitive((Number) cell.getNumericCellValue()));
                break;
              case BOOLEAN:
                values.put(names[cell.getColumnIndex()], new BooleanPrimitive((cell.getBooleanCellValue())));
                break;
              default:
                values.put(names[cell.getColumnIndex()], new StringPrimitive(cell.getStringCellValue()));
                break;
            }
          });

          response.addRow(values);
        }
      });
    }

    SCRIPT_LOGGER.debug("MSOfficeResource - readExcel -> Stop");
    return response;
  }
}
