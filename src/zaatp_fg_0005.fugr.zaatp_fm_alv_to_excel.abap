FUNCTION zaatp_fm_alv_to_excel.
*"----------------------------------------------------------------------
*"*"Local Interface:
*"  IMPORTING
*"     VALUE(IT_DATA) TYPE  ZAATP_TT_DATA
*"     VALUE(IT_INFO) TYPE  ZAATP_TT_INFO
*"     VALUE(IV_MAIL) TYPE  AD_SMTPADR OPTIONAL
*"     VALUE(IV_OPTION) TYPE  CHAR01 OPTIONAL
*"     VALUE(IV_TITLE) TYPE  SY-TITLE OPTIONAL
*"  EXPORTING
*"     VALUE(ET_RETURN) TYPE  BAPIRETTAB
*"----------------------------------------------------------------------
  DATA(lo_conv_excel) = NEW zcl_bc_alv_to_excel( ).

  DATA(lo_excel) = lo_conv_excel->alv_grid_to_abap2xslx(
                     it_data = it_data
                     it_info = it_info
                   ).


  TRY.
      zcl_bc_excel_output=>output(
        EXPORTING
          iv_option           = iv_option
          io_excel            = lo_excel
          iv_title            = iv_title
      ).
    CATCH zcx_excel INTO DATA(lx_excel).
      "to do: lx_excel-> bapirettaba yazılır.
  ENDTRY.


ENDFUNCTION.
