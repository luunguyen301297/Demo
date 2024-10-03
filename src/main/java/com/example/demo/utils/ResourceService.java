package com.viettel.solutions.resource.services;

import com.fasterxml.jackson.databind.ObjectMapper;
import com.viettel.solutions.resource.dtos.CredentialDTO;
import com.viettel.solutions.resource.dtos.request.AgencyReqDTO;
import com.viettel.solutions.resource.dtos.request.ResourceReqDTO;
import com.viettel.solutions.resource.dtos.response.ResourceResDTO;
import com.viettel.solutions.resource.repositories.AdministrativeUnitRepository;
import com.viettel.solutions.resource.repositories.CommonRepository;
import com.viettel.solutions.resource.repositories.ResourceRepository;
import com.viettel.solutions.resource.repositories.ResourceTypeRepository;
import com.viettel.solutions.resource.securities.ContextAwarePolicyEnforcement;
import com.viettel.solutions.resource.utils.ExcellUtils;
import com.viettel.solutions.resource.utils.ResponseUtils;
import com.viettel.solutions.resource.utils.Translator;
import com.viettel.solutions.resource.utils.exceptions.BusinessException;
import lombok.AllArgsConstructor;
import lombok.SneakyThrows;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbookType;
import org.springframework.core.env.Environment;
import org.springframework.http.HttpHeaders;
import org.springframework.http.HttpStatus;
import org.springframework.http.MediaType;
import org.springframework.http.ResponseEntity;
import org.springframework.stereotype.Component;
import org.springframework.stereotype.Service;
import org.springframework.web.multipart.MultipartFile;

import java.io.ByteArrayOutputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.math.BigInteger;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.util.*;

@Component
@Service
@AllArgsConstructor
public class ResourceService {
    private CredentialService credentialService;
    private ResourceRepository resourceRepository;

    private ContextAwarePolicyEnforcement policy;
    private CommonRepository commonRepository;
    private Environment env;
    private AdministrativeUnitRepository administrativeUnitRepository;
    private ResourceTypeRepository resourceTypeRepository;

    public ResponseEntity<Object> createResource(String ssoId, String tenant, ResourceReqDTO request){
        CredentialDTO credential = credentialService.getCredential(tenant, ssoId);

        if(!policy.checkPermission(credential.getUser(), request, "RESOURCE_CREATE"))
            return ResponseUtils.getResponseEntity(403, Translator.toLocale("common.message.access.denied", null), HttpStatus.FORBIDDEN);

        Map<String, Object> resource = resourceRepository.createResource(credential.getSchema(), credential.getUser(), request).get(0);

        return ResponseUtils.generateResponse(HttpStatus.CREATED, Translator.toLocale("common.message.action.successful", new String[]{"common.text.create", "common.text.resource"}), resource);
    }

    public ResponseEntity<Object> updateResource(String ssoId, String tenant, Long id, ResourceReqDTO request){
        CredentialDTO credential = credentialService.getCredential(tenant, ssoId);

        List<Map<String, Object>> items = resourceRepository.getResource(credential.getSchema(), id);

        if(items.size() == 0)
            throw new BusinessException(400, Translator.toLocale("common.message.not.found", new String[]{"common.text.resource"}));

        ResourceResDTO item = new ObjectMapper().convertValue(items.get(0), ResourceResDTO.class);


        if(!policy.checkPermission(credential.getUser(), item, "RESOURCE_UPDATE"))
            return ResponseUtils.getResponseEntity(403, Translator.toLocale("common.message.access.denied", null), HttpStatus.FORBIDDEN);

        resourceRepository.updateResource(credential.getSchema(), credential.getUser(), id, request);

        return ResponseUtils.getResponseEntity(200, Translator.toLocale("common.message.action.successful", new String[]{"common.text.edit", "common.text.resource"}), HttpStatus.OK);
    }

    public ResponseEntity<Object> deleteResource(String ssoId, String tenant, Long id){
        CredentialDTO credential = credentialService.getCredential(tenant, ssoId);

        List<Map<String, Object>> items = resourceRepository.getResource(credential.getSchema(), id);

        if(items.size() == 0)
            throw new BusinessException(400, Translator.toLocale("common.message.not.found", new String[]{"common.text.resource"}));

        ResourceResDTO item = new ObjectMapper().convertValue(items.get(0), ResourceResDTO.class);

        if(!policy.checkPermission(credential.getUser(), item, "RESOURCE_DELETE"))
            return ResponseUtils.getResponseEntity(403, Translator.toLocale("common.message.access.denied", null), HttpStatus.FORBIDDEN);

        resourceRepository.deleteResource(credential.getSchema(), credential.getUser().getUserId(), id);

        return ResponseUtils.getResponseEntity(200, Translator.toLocale("common.message.action.successful", new String[]{"common.text.delete", "common.text.resource"}), HttpStatus.OK);
    }

    public ResponseEntity<Object> getResource(String ssoId, String tenant, Long id){
        CredentialDTO credential = credentialService.getCredential(tenant, ssoId);

        List<Map<String, Object>> items = resourceRepository.getResource(credential.getSchema(), id);

        if(items.size() == 0)
            throw new BusinessException(400, Translator.toLocale("common.message.not.found", new String[]{"common.text.resource"}));

        ResourceResDTO resource = new ObjectMapper().convertValue(items.get(0), ResourceResDTO.class);

        if(!policy.checkPermission(credential.getUser(), resource, "RESOURCE_DETAIL"))
            return ResponseUtils.getResponseEntity(403, Translator.toLocale("common.message.access.denied", null), HttpStatus.FORBIDDEN);

        return ResponseUtils.generateResponse(HttpStatus.OK, "Successfully!", items.get(0));
    }

    public ResponseEntity<Object> getResources(String ssoId, String tenant, String name, String type, List<String> filterPlaces, String orderBy, Integer page, Integer size){
        CredentialDTO credential = credentialService.getCredential(tenant, ssoId);

        List<Map<String, Object>> resources = resourceRepository.getResources(credential.getSchema(), credential.getUser(), name, type, filterPlaces, orderBy, page, size);

        Integer totalElement = Integer.parseInt(resourceRepository.getTotalResources(credential.getSchema(), name, type, filterPlaces, credential.getUser()).get(0).get("total").toString());

        Integer totalPage = (int)Math.ceil((double)totalElement/size);

        return ResponseUtils.generateResponse(HttpStatus.OK, "Successfully!", resources, page, size, totalPage, totalElement);
    }

    public ResponseEntity<Object> importResources(String ssoId, String tenant, MultipartFile file){
        CredentialDTO credential = credentialService.getCredential(tenant, ssoId);
        ByteArrayOutputStream byteArrayOutputStream = new ByteArrayOutputStream();
        try{
            Workbook workbook = new XSSFWorkbook(file.getInputStream());
            Sheet sheet = workbook.getSheet(Translator.toLocale("common.template.resource.title", new String[]{}));
            int rowNumber = 5;
            List<Map<String, Object>> exception = new ArrayList<>();
            while (rowNumber < sheet.getLastRowNum()) {
                Row currentRow = sheet.getRow(rowNumber);
                ResourceReqDTO request = new ResourceReqDTO();
                int cellIdx = 0;
                while (cellIdx < currentRow.getLastCellNum()) {
                    Cell currentCell = currentRow.getCell(cellIdx);
                    DataFormatter formatter = new DataFormatter();
                    switch (cellIdx){
                        case 1:
                            if(formatter.formatCellValue(currentCell).equals(""))
                                break;
                            request.setName(currentCell.getStringCellValue());
                            break;
                        case 3:
                            try {
                                if(formatter.formatCellValue(currentCell).equals(""))
                                    break;
                                request.setResourceTypeId((long) currentCell.getNumericCellValue());
                                break;
                            } catch (Exception e) {
                                request.setResourceTypeId(null);
                            }
                            break;
                        case 4:
                            try {
                                if(formatter.formatCellValue(currentCell).equals(""))
                                    break;
                                request.setPhone(currentCell.getStringCellValue());
                            } catch (Exception e) {
                                return ResponseUtils.getResponseEntity(
                                        400, Translator.toLocale("common.message.warning.phone", new String[]{}), HttpStatus.OK);
                            }
                            break;
                        case 6:
                            try {
                                if(formatter.formatCellValue(currentCell).equals(""))
                                    break;
                                request.setProvinceCode(currentCell.getStringCellValue());
                                break;
                            } catch (Exception e) {
                                request.setProvinceCode(null);
                            }
                            break;
                        case 7:
                            if(formatter.formatCellValue(currentCell).equals(""))
                                break;
                            request.setLng(currentCell.getNumericCellValue());
                            break;
                        case 8:
                            if(formatter.formatCellValue(currentCell).equals(""))
                                break;
                            request.setLat(currentCell.getNumericCellValue());
                            break;
                        case 9:
                            if(formatter.formatCellValue(currentCell).equals(""))
                                break;
                            request.setAddress(currentCell.getStringCellValue());
                            break;
                        default:
                            break;
                    }
                    cellIdx++;
                }
                if(request.getName() != null) {
                    List<Map<String, Object>> resources = resourceRepository.createResourceFromExcel(credential.getSchema(), credential.getUser(), request);
                    if (resources != null && !resources.isEmpty() && resources.get(0) != null) {
                        Map<String, Object> resource = resources.get(0);
                        if (resource.get("code") != null) {
                            ExcellUtils.addException(exception, currentRow.getRowNum(), currentRow.getLastCellNum(), resource.get("message"));
                        } else if (resource.get("id") != null) {
                            ExcellUtils.addException(exception, currentRow.getRowNum(), currentRow.getLastCellNum(), Translator.toLocale("common.message.action.successful", new String[]{"common.text.create", "common.text.resource"}));
                        }
                    }
                }
                rowNumber++;
            }
            workbook.close();
            if (!exception.isEmpty()) {
                byteArrayOutputStream = ExcellUtils.checkExceptionExcel(exception, file);
                return new ResponseEntity<>(byteArrayOutputStream.toByteArray(), ExcellUtils.headers("Danh_sach_phong_ban_mau.xlsm"), HttpStatus.OK);
            }
            return ResponseUtils.getResponseEntity(
                    200, Translator.toLocale("common.message.action.successful", new String[]{"common.text.import", "common.text.user"}), HttpStatus.CREATED);
        } catch (Exception e){
            throw new BusinessException(400, Translator.toLocale("common.message.action.failed", new String[]{"importAgencies"}));
        }
    }

    @SneakyThrows
    public ResponseEntity<Object> downloadTemplateResource(String tenant, String ssoId) {
        CredentialDTO credential = credentialService.getCredential(tenant, ssoId);
        ByteArrayOutputStream byteArrayOutputStream = new ByteArrayOutputStream();

        try {
            Workbook wb = new XSSFWorkbook(XSSFWorkbookType.XLSM);
            XSSFSheet sheet1 = (XSSFSheet) wb.createSheet(Translator.toLocale("common.template.resource.title", new String[]{}));
            XSSFSheet sheet2 = (XSSFSheet) wb.createSheet("Sheet2");
            XSSFSheet sheet3 = (XSSFSheet) wb.createSheet("Sheet3");
            wb.setSheetVisibility(wb.getSheetIndex(sheet2), SheetVisibility.VERY_HIDDEN);
            wb.setSheetVisibility(wb.getSheetIndex(sheet3), SheetVisibility.VERY_HIDDEN);
            Map<String, CellStyle> styleMap = ExcellUtils.genStyleMap(wb);
            sheet1.createFreezePane(0, 5);

            int rowIdx = 0;
            Row row = sheet1.createRow(rowIdx);
            rowIdx++;
            row = sheet1.createRow(rowIdx);
            rowIdx++;
            Cell c = null;
            int cellIdx = 0;
            c = row.createCell(0);
            c.setCellValue(Translator.toLocale("common.template.resource.title", new String[]{}));
            c.setCellStyle(styleMap.get("titleCss"));
            sheet1.addMergedRegion(new CellRangeAddress(1, 1, 0, 9));
            row = sheet1.createRow(rowIdx);
            rowIdx++;

            //generate header row
            row = sheet1.createRow(rowIdx);
            rowIdx++;
            c = row.createCell(cellIdx);
            c.setCellValue(Translator.toLocale("common.text.stt", new String[]{}));
            c.setCellStyle(styleMap.get("colHeaderCss"));
            sheet1.setColumnWidth(cellIdx, 2000);
            cellIdx++;
            c = row.createCell(cellIdx);
            c.setCellValue(Translator.toLocale("common.template.resource.name", new String[]{}));
            c.setCellStyle(styleMap.get("colHeaderCss"));
            sheet1.setColumnWidth(cellIdx, 5000);
            cellIdx++;
            c = row.createCell(cellIdx);
            c.setCellValue(Translator.toLocale("common.template.resource.type", new String[]{}));
            c.setCellStyle(styleMap.get("colHeaderCss"));
            sheet1.setColumnWidth(cellIdx, 7000);
            cellIdx++;
            c = row.createCell(cellIdx);
            c.setCellStyle(styleMap.get("colHeaderCss"));
            sheet1.setColumnWidth(cellIdx, 7000);
            sheet1.setColumnHidden(cellIdx, true);
            int typeColHide = cellIdx;
            cellIdx++;
            c = row.createCell(cellIdx);
            c.setCellValue(Translator.toLocale("common.template.resource.phone", new String[]{}));
            c.setCellStyle(styleMap.get("colHeaderCss"));
            sheet1.setColumnWidth(cellIdx, 8000);
            cellIdx++;
            c = row.createCell(cellIdx);
            c.setCellValue(Translator.toLocale("common.template.resource.province", new String[]{}));
            c.setCellStyle(styleMap.get("colHeaderCss"));
            sheet1.setColumnWidth(cellIdx, 10000);
            cellIdx++;
            c = row.createCell(cellIdx);
            c.setCellStyle(styleMap.get("colHeaderCss"));
            sheet1.setColumnWidth(cellIdx, 10000);
            sheet1.setColumnHidden(cellIdx, true);
            int provinceColHide = cellIdx;
            cellIdx++;
            c = row.createCell(cellIdx);
            c.setCellValue(Translator.toLocale("common.template.resource.lgn", new String[]{}));
            c.setCellStyle(styleMap.get("colHeaderCss"));
            sheet1.setColumnWidth(cellIdx, 5000);
            cellIdx++;
            c = row.createCell(cellIdx);
            c.setCellValue(Translator.toLocale("common.template.resource.lat", new String[]{}));
            c.setCellStyle(styleMap.get("colHeaderCss"));
            sheet1.setColumnWidth(cellIdx, 5000);
            cellIdx++;
            c = row.createCell(cellIdx);
            c.setCellValue(Translator.toLocale("common.template.resource.address", new String[]{}));
            c.setCellStyle(styleMap.get("colHeaderCss"));
            sheet1.setColumnWidth(cellIdx, 8000);
            cellIdx++;

            cellIdx = 0;
            String[] colDescription = new String[]{
                    "",
                    Translator.toLocale("common.template.description.name", new String[]{}),
                    Translator.toLocale("common.template.description.type", new String[]{}), "",
                    Translator.toLocale("common.template.description.phone", new String[]{}),
                    Translator.toLocale("common.template.description.province", new String[]{}), "",
                    "", "",
                    Translator.toLocale("common.template.description.address", new String[]{})
            };
            ExcellUtils.generateCell(row, 3000, rowIdx, sheet1, cellIdx, c, styleMap.get("colDescriptionCss"), colDescription);
            rowIdx++;

            List<Map<String, Object>> resourceTypes = resourceTypeRepository.getResourceTypes(credential.getSchema(), credential.getUser(), "", "1,2,3", Arrays.asList("", "", ""), "", -1, 1);
            List<Map<String, Object>> config = commonRepository.configIOC(tenant);
            String schema = config.get(0).get("database").toString();
            String provinceCode = config.get(0).get("province_code").toString();
            List<Map<String, Object>> administrativeUnits = administrativeUnitRepository.getAdministrativeUnits(schema, provinceCode, "", Arrays.asList("", "", ""), "", -1, 1);

            //Lay du lieu va them vao sheet 2
            for (int i = 0; i < resourceTypes.size(); i++) {
                Map<String, Object> types = resourceTypes.get(i);
                row = sheet2.createRow(i);
                c = row.createCell(0);
                c.setCellValue((String) types.get("name"));
                c = row.createCell(1);
                c.setCellValue(((BigInteger) types.get("id")).doubleValue());
            }

            //Lay du lieu va them vao sheet 3
            for (int i = 0; i < administrativeUnits.size(); i++) {
                Map<String, Object> agency = administrativeUnits.get(i);
                row = sheet3.createRow(i);
                c = row.createCell(0);
                c.setCellValue((String) agency.get("name"));
                c = row.createCell(1);
                c.setCellValue((String) agency.get("code"));
            }
            //Lay danh sach ten don vi hanh chinh tu sheet 3 va them vao cot chon don vi o sheet 1
            ExcellUtils.generateSheetData(sheet1, rowIdx, rowIdx + 10, typeColHide - 1, typeColHide - 1, sheet2.getSheetName() + "!$A$1:$A$" + resourceTypes.size());

            //Lay danh sach ten phong ban tu sheet 2 va them vao cot chon vai tro o sheet 1
            ExcellUtils.generateSheetData(sheet1, rowIdx, rowIdx + 10, provinceColHide - 1, provinceColHide - 1, sheet3.getSheetName() + "!$A$1:$A$" + administrativeUnits.size());

            //Them code vao cell de lay ra id theo don vi/nhom nguoi dung
            for (int i = rowIdx; i < rowIdx + 10; i++) {
                row = sheet1.createRow(i);
                c = row.createCell(typeColHide);
                c.setCellFormula(ExcellUtils.formatVLookupCell("VLOOKUP(C%S, %S!$A$1:$B$%S, 2, FALSE)", String.valueOf((i + 1)), sheet2.getSheetName(), String.valueOf(resourceTypes.size())));
                c = row.createCell(provinceColHide);
                c.setCellFormula(ExcellUtils.formatVLookupCell("VLOOKUP(F%S, %S!$A$1:$B$%S, 2, FALSE)", String.valueOf((i + 1)), sheet3.getSheetName(), String.valueOf(administrativeUnits.size())));
            }
            ExcellUtils.setBorderStyle(sheet1, 5, 16, 0, 10, styleMap.get("normalCss"));
            wb.write(byteArrayOutputStream);
            wb.close();
        } catch (IOException e) {
            throw new RuntimeException(e);
        }

        return new ResponseEntity<>(byteArrayOutputStream.toByteArray(), ExcellUtils.headers("Danh_sach_tai_nguyen_mau.xlsm"), HttpStatus.OK);
    }
}
