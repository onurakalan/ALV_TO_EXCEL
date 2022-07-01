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
          iv_title            TYPE sy-title OPTIONAL
          iv_mail             TYPE ad_smtpadr OPTIONAL
        RAISING
          zcx_excel.

  PROTECTED SECTION.
  PRIVATE SECTION.
    DATA:
      xdata     TYPE xstring,             " Will be used for sending as email
      t_rawdata TYPE solix_tab,           " Will be used for downloading or open directly
      bytecount TYPE i,
      title TYPE sy-title,
      mail  TYPE ad_smtpadr.

    METHODS :
      download_frontend,
      _get_filename
        CHANGING
          cv_fullpath TYPE string
          cv_result   TYPE i,
      send_email,
      download_backend.
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
        lo_output->title = iv_title.
        lo_output->mail = iv_mail.


        CASE iv_option.
          WHEN zcl_bc_excel_outopt=>c_frontend.
            IF sy-batch IS INITIAL.
              lo_output->download_frontend( ).
            ELSE.
              MESSAGE e802(zabap2xlsx).
            ENDIF.

          WHEN zcl_bc_excel_outopt=>c_backend.
            lo_output->download_backend( ).

          WHEN zcl_bc_excel_outopt=>c_show.
            IF sy-batch IS INITIAL.
*          lo_output->display_online( ).
            ELSE.
              MESSAGE e803(zabap2xlsx).
            ENDIF.

          WHEN zcl_bc_excel_outopt=>c_mail.
            lo_output->send_email( ).

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

    CONCATENATE title '_' sy-datum '.XLSX' INTO lv_filename.

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


  METHOD send_email.

    DATA: bcs_exception        TYPE REF TO cx_bcs,
          errortext            TYPE string,
          cl_send_request      TYPE REF TO cl_bcs,
          cl_document          TYPE REF TO cl_document_bcs,
          cl_recipient         TYPE REF TO if_recipient_bcs,
          cl_sender            TYPE REF TO cl_cam_address_bcs,
          t_attachment_header  TYPE soli_tab,
          wa_attachment_header LIKE LINE OF t_attachment_header,
          attachment_subject   TYPE sood-objdes,

          sood_bytecount       TYPE sood-objlen,
          mail_title           TYPE so_obj_des,
          t_mailtext           TYPE soli_tab,
          wa_mailtext          LIKE LINE OF t_mailtext,
          send_to              TYPE adr6-smtp_addr,
          sent                 TYPE abap_bool,

          lv_filename          TYPE string.

    CONCATENATE title '_' sy-datum '.XLSX' INTO lv_filename.

    mail_title     = |{ title } Excel|.
    wa_mailtext    = 'Excel ektedir'.
    APPEND wa_mailtext TO t_mailtext.

    TRY.

        cl_send_request = cl_bcs=>create_persistent( ).
        cl_document = cl_document_bcs=>create_document( i_type    = 'RAW' "#EC NOTEXT
                                                        i_text    = t_mailtext
                                                        i_subject = mail_title ).

        attachment_subject  = lv_filename.
        CONCATENATE '&SO_FILENAME=' attachment_subject INTO wa_attachment_header.
        APPEND wa_attachment_header TO t_attachment_header.
* Attachment
        sood_bytecount = bytecount.
        cl_document->add_attachment(  i_attachment_type    = 'XLS' "#EC NOTEXT
                                      i_attachment_subject = attachment_subject
                                      i_attachment_size    = sood_bytecount
                                      i_att_content_hex    = t_rawdata
                                      i_attachment_header  = t_attachment_header ).

* add document to send request
        cl_send_request->set_document( cl_document ).

* add recipient(s) - here only 1 will be needed
        "kullanıcının maili..
        DATA : rpbenerr TYPE TABLE OF rpbenerr.

        CALL FUNCTION 'HR_FBN_GET_USER_EMAIL_ADDRESS'
          EXPORTING
            user_id       = sy-uname
            reaction      = 'X'
          IMPORTING
            email_address = send_to
          TABLES
            error_table   = rpbenerr
          .

*        send_to = 'aybuke.aydemir@improva.com.tr'.


        cl_recipient = cl_cam_address_bcs=>create_internet_address( send_to ).
        cl_send_request->add_recipient( cl_recipient ).

* Und abschicken
        sent = cl_send_request->send( i_with_error_screen = 'X' ).

        COMMIT WORK.

        IF sent = abap_true.
          MESSAGE s805(zabap2xlsx).
          MESSAGE 'Document ready to be sent - Check SOST or SCOT' TYPE 'I'.
        ELSE.
*          MESSAGE i804(zabap2xlsx) WITH p_email.
        ENDIF.

      CATCH cx_bcs INTO bcs_exception.
        errortext = bcs_exception->if_message~get_text( ).
        MESSAGE errortext TYPE 'I'.

    ENDTRY.
  ENDMETHOD.


  METHOD download_backend.
    DATA : lv_filename TYPE string.
    DATA: bytes_remain TYPE i.
    FIELD-SYMBOLS: <rawdata> LIKE LINE OF t_rawdata.

    CONCATENATE title '_' sy-datum '.XLSX' INTO lv_filename.

    OPEN DATASET lv_filename FOR OUTPUT IN BINARY MODE.
    CHECK sy-subrc = 0.

    bytes_remain = bytecount.

    LOOP AT t_rawdata ASSIGNING <rawdata>.

      AT LAST.
        CHECK bytes_remain >= 0.
        TRANSFER <rawdata> TO lv_filename LENGTH bytes_remain.
        EXIT.
      ENDAT.

      TRANSFER <rawdata> TO lv_filename.
      SUBTRACT 255 FROM bytes_remain.  " Solix has length 255

    ENDLOOP.

    CLOSE DATASET lv_filename.

    IF sy-repid <> sy-cprog AND sy-cprog IS NOT INITIAL.  " no need to display anything if download was selected and report was called for demo purposes
      LEAVE PROGRAM.
    ELSE.
      MESSAGE 'Data transferred to default backend directory' TYPE 'S'.
    ENDIF.
  ENDMETHOD.

ENDCLASS.
