"""
API Router for XML-Excel Conversion
Handles all conversion endpoints with proper error handling and validation
"""

from fastapi import APIRouter, UploadFile, File, HTTPException, Form
from fastapi.responses import FileResponse, JSONResponse
from typing import Optional, Literal
import os
import tempfile
import shutil
from pathlib import Path
import logging

from convertor.utils import ConversionService, ConversionError, ValidationError

# Configure logging
logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

# Create router
con_router = APIRouter(
    prefix="/api/converter",
    tags=["converter"]
)

# Initialize conversion service
conversion_service = ConversionService()


@con_router.post("/convert")
async def convert_file(
    file: UploadFile = File(..., description="File to convert (XML or Excel)"),
    direction: Optional[Literal["auto", "xml_to_excel", "excel_to_xml"]] = Form("auto"),
    root_tag: Optional[str] = Form(None, description="Custom root tag for XML output"),
    output_filename: Optional[str] = Form(None, description="Custom output filename")
):
    """
    Convert between XML and Excel formats.
    
    - **file**: Upload XML or Excel file
    - **direction**: Conversion direction (auto-detect by default)
    - **root_tag**: Custom root tag name for XML output (optional)
    - **output_filename**: Custom name for output file (optional)
    
    Returns the converted file for download.
    """
    temp_input = None
    temp_output = None
    
    try:
        # Validate file upload
        if not file.filename:
            raise HTTPException(
                status_code=400,
                detail="No file provided"
            )
        
        # Validate file extension
        file_ext = Path(file.filename).suffix.lower()
        allowed_extensions = {'.xml', '.xlsx', '.xls', '.xlsm'}
        
        if file_ext not in allowed_extensions:
            raise HTTPException(
                status_code=400,
                detail=f"Unsupported file type '{file_ext}'. Allowed: {', '.join(allowed_extensions)}"
            )
        
        # Create temporary directory for processing
        temp_dir = tempfile.mkdtemp(prefix="converter_")
        
        # Save uploaded file to temp location
        temp_input = os.path.join(temp_dir, file.filename)
        with open(temp_input, "wb") as buffer:
            shutil.copyfileobj(file.file, buffer)
        
        logger.info(f"Processing file: {file.filename} (size: {os.path.getsize(temp_input)} bytes)")
        
        # Determine conversion direction
        if direction == "auto":
            if file_ext == '.xml':
                direction = "xml_to_excel"
            elif file_ext in ['.xlsx', '.xls', '.xlsm']:
                direction = "excel_to_xml"
            else:
                raise HTTPException(
                    status_code=400,
                    detail="Cannot auto-detect conversion direction. Please specify explicitly."
                )
        
        # Perform conversion
        if direction == "xml_to_excel":
            temp_output = conversion_service.xml_to_excel(
                temp_input,
                output_path=None
            )
            media_type = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            default_ext = ".xlsx"
            
        elif direction == "excel_to_xml":
            temp_output = conversion_service.excel_to_xml(
                temp_input,
                output_path=None,
                root_tag=root_tag
            )
            media_type = "application/xml"
            default_ext = ".xml"
            
        else:
            raise HTTPException(
                status_code=400,
                detail=f"Invalid direction: {direction}"
            )
        
        # Determine output filename
        if output_filename:
            download_name = output_filename
            if not download_name.endswith(default_ext):
                download_name += default_ext
        else:
            original_stem = Path(file.filename).stem
            download_name = f"{original_stem}{default_ext}"
        
        logger.info(f"Conversion successful: {file.filename} -> {download_name}")
        
        # Return file for download
        return FileResponse(
            path=temp_output,
            media_type=media_type,
            filename=download_name,
            headers={
                "Content-Disposition": f"attachment; filename={download_name}"
            },
            background=cleanup_temp_files(temp_dir)
        )
        
    except ValidationError as e:
        logger.error(f"Validation error: {str(e)}")
        raise HTTPException(
            status_code=400,
            detail=f"Validation error: {str(e)}"
        )
        
    except ConversionError as e:
        logger.error(f"Conversion error: {str(e)}")
        raise HTTPException(
            status_code=422,
            detail=f"Conversion error: {str(e)}"
        )
        
    except Exception as e:
        logger.error(f"Unexpected error: {str(e)}", exc_info=True)
        # Cleanup on error
        if temp_input and os.path.exists(os.path.dirname(temp_input)):
            shutil.rmtree(os.path.dirname(temp_input), ignore_errors=True)
        
        raise HTTPException(
            status_code=500,
            detail=f"Internal server error: {str(e)}"
        )


@con_router.post("/validate")
async def validate_file(
    file: UploadFile = File(..., description="File to validate")
):
    """
    Validate if a file can be converted.
    
    Returns file information and conversion compatibility.
    """
    try:
        if not file.filename:
            raise HTTPException(status_code=400, detail="No file provided")
        
        file_ext = Path(file.filename).suffix.lower()
        file_size = 0
        
        # Read file to get size
        content = await file.read()
        file_size = len(content)
        await file.seek(0)  # Reset file pointer
        
        # Validate extension
        is_xml = file_ext == '.xml'
        is_excel = file_ext in ['.xlsx', '.xls', '.xlsm']
        
        if not (is_xml or is_excel):
            return JSONResponse(
                status_code=200,
                content={
                    "valid": False,
                    "filename": file.filename,
                    "size": file_size,
                    "error": f"Unsupported file type: {file_ext}",
                    "supported_types": [".xml", ".xlsx", ".xls", ".xlsm"]
                }
            )
        
        # Determine conversion options
        can_convert_to = []
        if is_xml:
            can_convert_to.append("excel")
        if is_excel:
            can_convert_to.append("xml")
        
        return JSONResponse(
            status_code=200,
            content={
                "valid": True,
                "filename": file.filename,
                "extension": file_ext,
                "size": file_size,
                "size_mb": round(file_size / (1024 * 1024), 2),
                "file_type": "xml" if is_xml else "excel",
                "can_convert_to": can_convert_to
            }
        )
        
    except Exception as e:
        logger.error(f"Validation error: {str(e)}")
        raise HTTPException(
            status_code=500,
            detail=f"Validation failed: {str(e)}"
        )


@con_router.get("/health")
async def health_check():
    """
    Health check endpoint.
    """
    return {
        "status": "healthy",
        "service": "xml-excel-converter",
        "version": "1.0.0"
    }


def cleanup_temp_files(temp_dir: str):
    """
    Background task to cleanup temporary files after response is sent.
    """
    async def _cleanup():
        try:
            if os.path.exists(temp_dir):
                shutil.rmtree(temp_dir, ignore_errors=True)
                logger.info(f"Cleaned up temp directory: {temp_dir}")
        except Exception as e:
            logger.error(f"Error cleaning up temp files: {str(e)}")
    
    return _cleanup