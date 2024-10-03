package com.viettel.solutions.resource.controllers;

import com.viettel.solutions.resource.dtos.param.ResourceParamDTO;
import com.viettel.solutions.resource.dtos.request.ResourceReqDTO;
import com.viettel.solutions.resource.securities.ContextAwarePolicyEnforcement;
import com.viettel.solutions.resource.services.ExcelService;
import com.viettel.solutions.resource.services.ResourceOnMapService;
import com.viettel.solutions.resource.services.ResourceService;
import com.viettel.solutions.resource.utils.ResponseUtils;
import com.viettel.solutions.resource.utils.Translator;
import com.viettel.solutions.resource.utils.exceptions.BusinessException;
import com.viettel.solutions.resource.utils.validates.CreateValidator;
import io.swagger.v3.oas.annotations.Operation;
import io.swagger.v3.oas.annotations.media.Content;
import io.swagger.v3.oas.annotations.media.Schema;
import io.swagger.v3.oas.annotations.responses.ApiResponse;
import io.swagger.v3.oas.annotations.responses.ApiResponses;
import io.swagger.v3.oas.annotations.security.SecurityRequirement;
import lombok.AllArgsConstructor;
import org.springframework.http.HttpStatus;
import org.springframework.http.ResponseEntity;
import org.springframework.security.core.Authentication;
import org.springframework.security.core.context.SecurityContextHolder;
import org.springframework.validation.annotation.Validated;
import org.springframework.web.bind.annotation.*;
import org.springframework.web.multipart.MultipartFile;

import javax.annotation.security.RolesAllowed;
import java.util.List;

@RestController
@AllArgsConstructor
@CrossOrigin(origins = {"*"})
@RequestMapping("/resources")
@Validated
public class ResourceController {

    private ResourceService resourceService;
    private CreateValidator createValidator;

    @Operation(
            summary = "Create a new resource",
            description = "Tạo mới một tài nguyên")
    @SecurityRequirement(name = "Bearer Authentication")
    @ApiResponses({
            @ApiResponse(responseCode = "200", content = { @Content(schema = @Schema(), mediaType = "application/json") }),
            @ApiResponse(responseCode = "404", content = { @Content(schema = @Schema()) }),
            @ApiResponse(responseCode = "500", content = { @Content(schema = @Schema()) }) })
    @PostMapping
//    @RolesAllowed({"LEADER"})
    public ResponseEntity<Object> createResource(@RequestHeader(value = "tenant", required = true) String tenant,
                                                 @RequestBody ResourceReqDTO request){
        try {
            String ssoId = SecurityContextHolder.getContext().getAuthentication().getName();
            ResourceReqDTO validate = createValidator.validateResourceReq(request);
            ResponseEntity<Object> response = resourceService.createResource(ssoId, tenant, validate);
            return response;
        } catch (Exception e){
            if(e instanceof BusinessException)
                return ResponseUtils.getResponseEntity(((BusinessException) e).getCodeError(), e.getMessage(), HttpStatus.BAD_REQUEST);
            return ResponseUtils.getResponseEntity(400, e.getMessage(), HttpStatus.BAD_REQUEST);
        }
    }

    @Operation(
            summary = "Update a specific resource",
            description = "Cập nhật một tài nguyên cụ thể")
    @SecurityRequirement(name = "Bearer Authentication")
    @ApiResponses({
            @ApiResponse(responseCode = "200", content = { @Content(schema = @Schema(), mediaType = "application/json") }),
            @ApiResponse(responseCode = "404", content = { @Content(schema = @Schema()) }),
            @ApiResponse(responseCode = "500", content = { @Content(schema = @Schema()) }) })
    @PutMapping("{id}")
//    @RolesAllowed({"LEADER"})
    public ResponseEntity<Object> updateResource(@RequestHeader(value = "tenant", required = true) String tenant,
                                                 @PathVariable Long id,
                                                 @RequestBody ResourceReqDTO request){
        try {
            String ssoId = SecurityContextHolder.getContext().getAuthentication().getName();
            ResourceReqDTO validate = createValidator.validateResourceReq(request);
            ResponseEntity<Object> response = resourceService.updateResource(ssoId, tenant, id, validate);
            return response;
        } catch (Exception e){
            if(e instanceof BusinessException)
                return ResponseUtils.getResponseEntity(((BusinessException) e).getCodeError(), e.getMessage(), HttpStatus.BAD_REQUEST);
            return ResponseUtils.getResponseEntity(400, e.getMessage(), HttpStatus.BAD_REQUEST);
        }
    }

    @Operation(
            summary = "Delete a specific resource",
            description = "Xóa một tài nguyên cụ thể")
    @SecurityRequirement(name = "Bearer Authentication")
    @ApiResponses({
            @ApiResponse(responseCode = "200", content = { @Content(schema = @Schema(), mediaType = "application/json") }),
            @ApiResponse(responseCode = "404", content = { @Content(schema = @Schema()) }),
            @ApiResponse(responseCode = "500", content = { @Content(schema = @Schema()) }) })
    @DeleteMapping("{id}")
//    @RolesAllowed({"LEADER"})
    public ResponseEntity<Object> deleteResource(@RequestHeader(value = "tenant", required = true) String tenant,
                                                 @PathVariable Long id){
        try {
            String ssoId = SecurityContextHolder.getContext().getAuthentication().getName();
            ResponseEntity<Object> response = resourceService.deleteResource(ssoId, tenant, id);
            return response;
        } catch (Exception e){
            if(e instanceof BusinessException)
                return ResponseUtils.getResponseEntity(((BusinessException) e).getCodeError(), e.getMessage(), HttpStatus.BAD_REQUEST);
            return ResponseUtils.getResponseEntity(400, e.getMessage(), HttpStatus.BAD_REQUEST);
        }
    }

    @Operation(
            summary = "Retrieve a specific resource",
            description = "Lấy thông tin một tài nguyên cụ thể")
    @SecurityRequirement(name = "Bearer Authentication")
    @ApiResponses({
            @ApiResponse(responseCode = "200", content = { @Content(schema = @Schema(), mediaType = "application/json") }),
            @ApiResponse(responseCode = "404", content = { @Content(schema = @Schema()) }),
            @ApiResponse(responseCode = "500", content = { @Content(schema = @Schema()) }) })
    @GetMapping("{id}")
//    @RolesAllowed({"LEADER"})
    public ResponseEntity<Object> getResource(@RequestHeader(value = "tenant", required = true) String tenant,
                                              @PathVariable Long id){
        try {
            String ssoId = SecurityContextHolder.getContext().getAuthentication().getName();
            ResponseEntity<Object> response = resourceService.getResource(ssoId, tenant, id);
            return response;
        } catch (Exception e){
            if(e instanceof BusinessException)
                return ResponseUtils.getResponseEntity(((BusinessException) e).getCodeError(), e.getMessage(), HttpStatus.BAD_REQUEST);
            return ResponseUtils.getResponseEntity(400, e.getMessage(), HttpStatus.BAD_REQUEST);
        }
    }

    @Operation(
            summary = "Retrieve resource list",
            description = "Lấy danh sách tài nguyên")
    @SecurityRequirement(name = "Bearer Authentication")
    @ApiResponses({
            @ApiResponse(responseCode = "200", content = { @Content(schema = @Schema(), mediaType = "application/json") }),
            @ApiResponse(responseCode = "404", content = { @Content(schema = @Schema()) }),
            @ApiResponse(responseCode = "500", content = { @Content(schema = @Schema()) }) })
    @GetMapping
    public ResponseEntity<Object> getResources(@RequestHeader(value = "tenant") String tenant,
                                               @RequestParam(value = "name", defaultValue = "", required = false) String name,
                                               @RequestParam(value = "type", defaultValue = "", required = false) String type,
                                               @RequestParam(value = "place", defaultValue = ",,", required = false) List<String> filterPlaces,
                                               @RequestParam(value = "orderBy", defaultValue = "", required = false) String orderBy,
                                               @RequestParam(value = "page", defaultValue = "-1", required = false) Integer page,
                                               @RequestParam(value = "size", defaultValue = "1", required = false) Integer size){
        try {
            String ssoId = SecurityContextHolder.getContext().getAuthentication().getName();
            ResourceParamDTO validate = createValidator.validateResourceParam(name, type, filterPlaces, orderBy, page, size);
            ResponseEntity<Object> response = resourceService.getResources(
                    ssoId, tenant, validate.getName(), validate.getType(), validate.getPlace(), validate.getOrderBy(), validate.getPage(), validate.getSize());
            return response;
        } catch (Exception e){
            if(e instanceof BusinessException)
                return ResponseUtils.getResponseEntity(((BusinessException) e).getCodeError(), e.getMessage(), HttpStatus.BAD_REQUEST);
            return ResponseUtils.getResponseEntity(400, e.getMessage(), HttpStatus.BAD_REQUEST);
        }
    }

    @Operation(
            summary = "Download biểu mẫu tài nguyên",
            description = "")
    @SecurityRequirement(name = "Bearer Authentication")
    @ApiResponses({
            @ApiResponse(responseCode = "200", content = { @Content(schema = @Schema(), mediaType = "application/json") }),
            @ApiResponse(responseCode = "404", content = { @Content(schema = @Schema()) }),
            @ApiResponse(responseCode = "500", content = { @Content(schema = @Schema()) }) })
    @GetMapping(path = "download-resource-template")
    public ResponseEntity<Object> downloadTemplateResource(@RequestHeader(value = "tenant") String tenant){
        try {
            String userId = SecurityContextHolder.getContext().getAuthentication().getName();
            ResponseEntity<Object> response = resourceService.downloadTemplateResource(tenant, userId);
            return response;
        } catch (Exception e){
            if(e instanceof BusinessException)
                return ResponseUtils.getResponseEntity(((BusinessException) e).getCodeError(), e.getMessage(), HttpStatus.BAD_REQUEST);
            return ResponseUtils.getResponseEntity(400, e.getMessage(), HttpStatus.BAD_REQUEST);
        }
    }

    @Operation(
            summary = "Import danh sách tài nguyên từ file excel",
            description = "")
    @SecurityRequirement(name = "Bearer Authentication")
    @ApiResponses({
            @ApiResponse(responseCode = "200", content = { @Content(schema = @Schema(), mediaType = "application/json") }),
            @ApiResponse(responseCode = "404", content = { @Content(schema = @Schema()) }),
            @ApiResponse(responseCode = "500", content = { @Content(schema = @Schema()) }) })
    @PostMapping(path = "import-resources", consumes = {"multipart/form-data"})
    public ResponseEntity<Object> importResources(@RequestHeader(value = "tenant") String tenant,
                                                 @RequestParam("file") MultipartFile file){
        try {
            if (ExcelService.hasExcelFormat(file)) {
                String userId = SecurityContextHolder.getContext().getAuthentication().getName();
                ResponseEntity<Object> response = resourceService.importResources(userId, tenant, file);
                return response;
            }
            return ResponseUtils.getResponseEntity(
                    400, Translator.toLocale("common.message.action.failed", new String[]{"common.text.import"}), HttpStatus.OK);
        } catch (Exception e){
            if(e instanceof BusinessException)
                return ResponseUtils.getResponseEntity(((BusinessException) e).getCodeError(), e.getMessage(), HttpStatus.BAD_REQUEST);
            return ResponseUtils.getResponseEntity(400, e.getMessage(), HttpStatus.BAD_REQUEST);
        }
    }
}
