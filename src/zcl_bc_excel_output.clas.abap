CLASS zcl_bc_excel_output DEFINITION
  PUBLIC
  FINAL
  CREATE PRIVATE .

  PUBLIC SECTION.
    CLASS-METHODS:
      output
        IMPORTING
          iv_option           TYPE char01
          io_excel            TYPE REF TO zcl_excel
          iv_writerclass_name TYPE clike OPTIONAL
          iv_info_message     TYPE abap_bool DEFAULT abap_true
        RAISING
          zcx_excel.

  PROTECTED SECTION.
  PRIVATE SECTION.
    DATA:
      xdata     TYPE xstring,             " Will be used for sending as email
      t_rawdata TYPE solix_tab,           " Will be used for downloading or open directly
      bytecount TYPE i.

    METHODS :
      download_frontend,
      _get_filename
        CHANGING
          cv_fullpath TYPE string
          cv_result   TYPE i.
ENDCLASS.



CLASS zcl_bc_excel_output IMPLEMENTATION.
  METHOD output.
    DATA: lo_output TYPE REF TO zcl_bc_excel_output,
          lo_writer TYPE REF TO zif_excel_writer,
          lo_error  TYPE REF TO zcx_excel..

    TRY.
        IF iv_writerclass_name IS INITIAL.
          CREATE OBJECT lo_output.
          CREATE OBJECT lo_writer TYPE zcl_excel_writer_2007.
        ELSE.
          CREATE OBJECT lo_output.
          CREATE OBJECT lo_writer TYPE (iv_writerclass_name).
        ENDIF.

        lo_output->xdata = lo_writer->write_file( io_excel ).
        lo_output->t_rawdata = cl_bcs_convert=>xstring_to_solix( iv_xstring  = lo_output->xdata ).
        lo_output->bytecount = xstrlen( lo_output->xdata ).


        CASE iv_option.
          WHEN zcl_bc_excel_outopt=>c_frontend.
            IF sy-batch IS INITIAL.
              lo_output->download_frontend( ).
            ELSE.
              MESSAGE e802(zabap2xlsx).
            ENDIF.

          WHEN zcl_bc_excel_outopt=>c_backend.
*        lo_output->download_backend( ).

          WHEN zcl_bc_excel_outopt=>c_show.
            IF sy-batch IS INITIAL.
*          lo_output->display_online( ).
            ELSE.
              MESSAGE e803(zabap2xlsx).
            ENDIF.

          WHEN zcl_bc_excel_outopt=>c_mail.
*        cl_output->send_email( ).

        ENDCASE.

      CATCH zcx_excel INTO DATA(cl_error).
        IF iv_info_message = abap_true.
          MESSAGE cl_error TYPE 'I' DISPLAY LIKE 'E'.
        ELSE.
          RAISE EXCEPTION cl_error.
        ENDIF.

    ENDTRY.


  ENDMETHOD.

  METHOD _get_filename.
    DATA : lv_path     TYPE string,
           lv_filename TYPE string.

    cl_gui_frontend_services=>get_desktop_directory( CHANGING desktop_directory = lv_path ).
    cl_gui_cfw=>flush( ).

    CONCATENATE sy-title '_' sy-datum '.XLSX' INTO lv_filename.

    CALL METHOD cl_gui_frontend_services=>file_save_dialog
      EXPORTING
        window_title      = 'Export Excel'
        default_extension = 'XLSX'
        file_filter       = 'Excel dosyası (*.XLSX)'
        default_file_name = lv_filename
        initial_directory = lv_path
      CHANGING
        filename          = lv_filename
        path              = lv_path
        fullpath          = cv_fullpath
        user_action       = cv_result.

    IF cv_fullpath CA '/'.
      REPLACE REGEX '([^/])\s*$' IN cv_fullpath WITH '$1/' .
    ELSE.
      REPLACE REGEX '([^\\])\s*$' IN cv_fullpath WITH '$1\\'.
    ENDIF.

  ENDMETHOD.

  METHOD download_frontend.
    DATA : lv_fullpath TYPE string,
           lv_result   TYPE i,
           lv_message  TYPE string.

    me->_get_filename( CHANGING cv_fullpath = lv_fullpath
                                cv_result   = lv_result ).

    cl_gui_frontend_services=>gui_download(
      EXPORTING
        bin_filesize              = bytecount
        filename                  = lv_fullpath
        filetype                  = 'BIN'
        no_auth_check             = abap_true
      CHANGING
        data_tab                  = t_rawdata
      EXCEPTIONS
        file_write_error          = 1
        no_batch                  = 2
        gui_refuse_filetransfer   = 3
        invalid_type              = 4
        no_authority              = 5
        unknown_error             = 6
        header_not_allowed        = 7
        separator_not_allowed     = 8
        filesize_not_allowed      = 9
        header_too_long           = 10
        dp_error_create           = 11
        dp_error_send             = 12
        dp_error_write            = 13
        unknown_dp_error          = 14
        access_denied             = 15
        dp_out_of_memory          = 16
        disk_full                 = 17
        dp_timeout                = 18
        file_not_found            = 19
        dataprovider_exception    = 20
        control_flush_error       = 21
        not_supported_by_gui      = 22
        error_no_gui              = 23
        OTHERS                    = 24
    ).
    IF sy-subrc <> 0.
      MESSAGE ID sy-msgid TYPE sy-msgty NUMBER sy-msgno
              WITH sy-msgv1 sy-msgv2 sy-msgv3 sy-msgv4 INTO lv_message.

      RAISE EXCEPTION TYPE zcx_excel EXPORTING error = lv_message.

    ENDIF.

  ENDMETHOD.

ENDCLASS.
