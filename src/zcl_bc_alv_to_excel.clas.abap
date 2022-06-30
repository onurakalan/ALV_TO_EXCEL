**********************************************************************
* Abap : onur.akalan@improva.com.tr
* Date : 30.06.2022
**********************************************************************
CLASS zcl_bc_alv_to_excel DEFINITION
  PUBLIC
  FINAL
  CREATE PUBLIC .

  PUBLIC SECTION.

    METHODS:
      constructor,
      alv_grid_to_abap2xslx
        IMPORTING
          ir_grid         TYPE REF TO cl_gui_alv_grid
        RETURNING
          VALUE(ro_excel) TYPE REF TO zcl_excel.


  PROTECTED SECTION.
  PRIVATE SECTION.
    CONSTANTS:
      c_red_cell   TYPE lvc_istyle VALUE '7',
      c_green_cell TYPE lvc_istyle VALUE '6'.
      "to do : add other color

    CONSTANTS:
      c_subtotal TYPE lvc_istyle VALUE '36',
      c_total    TYPE lvc_istyle VALUE '44'.

    CONSTANTS:
      c_header_count TYPE i VALUE 1. "Excel Header Line Count

    TYPES:
      tt_data TYPE SORTED TABLE OF lvc_s_data WITH NON-UNIQUE KEY row_pos col_pos,
      tt_info TYPE SORTED TABLE OF lvc_s_info WITH NON-UNIQUE KEY col_pos.

    DATA :
      _mv_style_title              TYPE zexcel_cell_style,
      _mv_style_total              TYPE zexcel_cell_style,
      _mv_style_subtotal_sides     TYPE zexcel_cell_style,
      _mv_style_subtotal_down      TYPE zexcel_cell_style,
      _mv_style_subtotal_topbottom TYPE zexcel_cell_style,
      _mv_style_green_cell         TYPE zexcel_cell_style,
      _mv_style_red_cell           TYPE zexcel_cell_style,
      _mv_style_cell               TYPE zexcel_cell_style.

    DATA :
        _mv_sub_begline TYPE i.

    METHODS :
      _get_alv_data
        IMPORTING
          ir_grid TYPE REF TO cl_gui_alv_grid
        EXPORTING
          et_data TYPE tt_data
          et_info TYPE tt_info,
      _set_style
        IMPORTING
          io_excel TYPE REF TO zcl_excel,
      _get_sorting_column
        IMPORTING
          it_info    TYPE tt_info
        EXPORTING
          ev_colpos1 TYPE i
          ev_colpos2 TYPE i
          ev_colpos3 TYPE i
          ev_colpos4 TYPE i
          ev_colpos5 TYPE i
          ev_colpos6 TYPE i
          ev_colpos7 TYPE i
          ev_colpos8 TYPE i,
      _data_to_excel
        EXPORTING
          it_data  TYPE zcl_bc_alv_to_excel=>tt_data
          it_info  TYPE zcl_bc_alv_to_excel=>tt_info
        CHANGING
          io_excel TYPE REF TO zcl_excel.


    CLASS-METHODS:
      _progress_indicator
        IMPORTING
          iv_text       TYPE any
          iv_percentage TYPE any.


ENDCLASS.



CLASS zcl_bc_alv_to_excel IMPLEMENTATION.

  METHOD constructor.
    _mv_sub_begline = c_header_count + 1.
  ENDMETHOD.

  METHOD alv_grid_to_abap2xslx.
    DATA:
      lo_excel     TYPE REF TO zcl_excel.

    "GET ALV DATA
    """"""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""
    me->_get_alv_data(
      EXPORTING
        ir_grid = ir_grid
      IMPORTING
        et_data = DATA(lt_data)
        et_info = DATA(lt_info)
    ).

    "CREATE EXCEL
    """"""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""
    CREATE OBJECT lo_excel.

    "SET STYLE
    """""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""
    me->_set_style( io_excel = lo_excel ).

    "DATA TO EXCEL
    """"""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""
    me->_data_to_excel(
        IMPORTING
            it_data = lt_data
            it_info = lt_info
         CHANGING
            io_excel = lo_excel ).
  ENDMETHOD.

  METHOD _get_alv_data.
    DATA :
      lr_grid_facade TYPE REF TO cl_salv_gui_grid_facade.

    FIELD-SYMBOLS:
      <lt_data> TYPE lvc_t_data,
      <lt_info> TYPE lvc_t_info.

    _progress_indicator(
        iv_text       = TEXT-001
        iv_percentage = 0
    ).
    """"""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""
    CREATE OBJECT lr_grid_facade
      EXPORTING
        o_grid = ir_grid.


    lr_grid_facade->if_salv_gui_grid_lvc_data~get_all_data(
      EXPORTING
        gui_type      = cl_salv_gru_view_grid=>c_gui_type-windows
        view          = if_salv_c_function=>view_excel
      IMPORTING
        table_lines   = DATA(lt_table_lines)
        rt_data       = DATA(lr_data)
        rt_info       = DATA(lr_info)
        rt_idpo       = DATA(lr_idpo)
        rt_poid       = DATA(lr_poid)
        rt_roid       = DATA(lr_roid)
        t_start_index = DATA(lr_start_index)
    ).

    ASSIGN lr_data->* TO <lt_data>.
    ASSIGN lr_info->* TO <lt_info>.

    INSERT LINES OF <lt_data> INTO TABLE et_data.
    INSERT LINES OF <lt_info> INTO TABLE et_info.

    """"""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""
    _progress_indicator(
        iv_text       = TEXT-001
        iv_percentage = 100
    ).

  ENDMETHOD.

  METHOD _progress_indicator.
    CALL FUNCTION 'SAPGUI_PROGRESS_INDICATOR'
      EXPORTING
        percentage = iv_percentage
        text       = |{ iv_text }... %{ iv_percentage }|.
  ENDMETHOD.

  METHOD _set_style.
    "header line
    """"""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""
    DATA(lo_title_style) = io_excel->add_new_style( ).
    lo_title_style->fill->filltype     = zcl_excel_style_fill=>c_fill_solid.
    lo_title_style->fill->fgcolor-rgb  = zcl_excel_style_color=>c_gray.
    lo_title_style->font->bold = abap_true.
    lo_title_style->font->size = 10.
    lo_title_style->alignment->wraptext = 'X'.
    lo_title_style->alignment->horizontal = zcl_excel_style_alignment=>c_horizontal_center.
    lo_title_style->alignment->vertical = zcl_excel_style_alignment=>c_vertical_center.
    CREATE OBJECT lo_title_style->borders->allborders.
    lo_title_style->borders->allborders->border_style = zcl_excel_style_border=>c_border_thin.
    _mv_style_title = lo_title_style->get_guid( ).

    "total line style
    """"""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""
    DATA(lo_total_style) = io_excel->add_new_style( ).
    lo_total_style->fill->filltype     = zcl_excel_style_fill=>c_fill_solid.
    lo_total_style->fill->fgcolor-rgb  = zcl_excel_style_color=>c_darkyellow.
    lo_total_style->font->bold = abap_true.
    lo_total_style->font->size = 10.
    lo_total_style->alignment->wraptext = 'X'.
    lo_total_style->alignment->horizontal = zcl_excel_style_alignment=>c_horizontal_center.
    lo_total_style->alignment->vertical = zcl_excel_style_alignment=>c_vertical_center.
    CREATE OBJECT lo_total_style->borders->allborders.
    lo_total_style->borders->allborders->border_style = zcl_excel_style_border=>c_border_thin.
    _mv_style_total = lo_total_style->get_guid( ).

    "subtotal style - sides border
    """"""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""
    DATA(lo_subtotal_style1) = io_excel->add_new_style( ).
    lo_subtotal_style1->fill->filltype     = zcl_excel_style_fill=>c_fill_solid.
    lo_subtotal_style1->fill->fgcolor-rgb  = zcl_excel_style_color=>c_yellow.
    lo_subtotal_style1->font->bold = abap_true.
    lo_subtotal_style1->font->size = 10.
    lo_subtotal_style1->alignment->wraptext = 'X'.
    lo_subtotal_style1->alignment->horizontal = zcl_excel_style_alignment=>c_horizontal_center.
    lo_subtotal_style1->alignment->vertical = zcl_excel_style_alignment=>c_vertical_center.
    CREATE OBJECT lo_subtotal_style1->borders->left.
    CREATE OBJECT lo_subtotal_style1->borders->right.
    lo_subtotal_style1->borders->left->border_style = zcl_excel_style_border=>c_border_thin.
    lo_subtotal_style1->borders->right->border_style = zcl_excel_style_border=>c_border_thin.
    _mv_style_subtotal_sides = lo_subtotal_style1->get_guid( ).

    "subtotal style - down border
    """"""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""
    DATA(lo_subtotal_style2) = io_excel->add_new_style( ).
    lo_subtotal_style2->fill->filltype     = zcl_excel_style_fill=>c_fill_solid.
    lo_subtotal_style2->fill->fgcolor-rgb  = zcl_excel_style_color=>c_yellow.
    lo_subtotal_style2->font->bold = abap_true.
    lo_subtotal_style2->font->size = 10.
    lo_subtotal_style2->alignment->wraptext = 'X'.
    lo_subtotal_style2->alignment->horizontal = zcl_excel_style_alignment=>c_horizontal_center.
    lo_subtotal_style2->alignment->vertical = zcl_excel_style_alignment=>c_vertical_center.
    CREATE OBJECT lo_subtotal_style2->borders->down.
    lo_subtotal_style2->borders->down->border_style = zcl_excel_style_border=>c_border_thin.
    _mv_style_subtotal_down = lo_subtotal_style2->get_guid( ).

    "subtotal style - top down border
    """"""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""
    DATA(lo_subtotal_style3) = io_excel->add_new_style( ).
    lo_subtotal_style3->fill->filltype     = zcl_excel_style_fill=>c_fill_solid.
    lo_subtotal_style3->fill->fgcolor-rgb  = zcl_excel_style_color=>c_yellow.
    lo_subtotal_style3->font->bold = abap_true.
    lo_subtotal_style3->font->size = 10.
    lo_subtotal_style3->alignment->wraptext = 'X'.
    lo_subtotal_style3->alignment->horizontal = zcl_excel_style_alignment=>c_horizontal_center.
    lo_subtotal_style3->alignment->vertical = zcl_excel_style_alignment=>c_vertical_center.
    CREATE OBJECT lo_subtotal_style3->borders->top.
    CREATE OBJECT lo_subtotal_style3->borders->down.
    lo_subtotal_style3->borders->top->border_style = zcl_excel_style_border=>c_border_thin.
    lo_subtotal_style3->borders->down->border_style = zcl_excel_style_border=>c_border_thin.
    _mv_style_subtotal_topbottom = lo_subtotal_style3->get_guid( ).

    "cell - green
    """"""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""
    DATA(lo_green_style) = io_excel->add_new_style( ).
    lo_green_style->fill->filltype     = zcl_excel_style_fill=>c_fill_solid.
    lo_green_style->fill->fgcolor-rgb  = zcl_excel_style_color=>c_green.
    lo_green_style->font->size = 10.
    lo_green_style->alignment->wraptext = 'X'.
    lo_green_style->alignment->horizontal = zcl_excel_style_alignment=>c_horizontal_center.
    lo_green_style->alignment->vertical = zcl_excel_style_alignment=>c_vertical_center.
    CREATE OBJECT lo_green_style->borders->allborders.
    lo_green_style->borders->allborders->border_style = zcl_excel_style_border=>c_border_thin.
    _mv_style_green_cell = lo_green_style->get_guid( ).

    "cell - red
    """"""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""
    DATA(lo_red_style) = io_excel->add_new_style( ).
    lo_red_style->fill->filltype     = zcl_excel_style_fill=>c_fill_solid.
    lo_red_style->fill->fgcolor-rgb  = zcl_excel_style_color=>c_red.
    lo_red_style->font->size = 10.
    lo_red_style->alignment->wraptext = 'X'.
    lo_red_style->alignment->horizontal = zcl_excel_style_alignment=>c_horizontal_center.
    lo_red_style->alignment->vertical = zcl_excel_style_alignment=>c_vertical_center.
    CREATE OBJECT lo_red_style->borders->allborders.
    lo_red_style->borders->allborders->border_style = zcl_excel_style_border=>c_border_thin.
    _mv_style_red_cell = lo_red_style->get_guid( ).

    "cell others
    """"""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""
    DATA(lo_cell_style) = io_excel->add_new_style( ).
    lo_cell_style->fill->filltype     = zcl_excel_style_fill=>c_fill_solid.
    lo_cell_style->fill->fgcolor-rgb  = zcl_excel_style_color=>c_white.
    lo_cell_style->font->size = 10.
    lo_cell_style->alignment->wraptext = 'X'.
    lo_cell_style->alignment->horizontal = zcl_excel_style_alignment=>c_horizontal_center.
    lo_cell_style->alignment->vertical = zcl_excel_style_alignment=>c_vertical_center.
    CREATE OBJECT lo_cell_style->borders->allborders.
    lo_cell_style->borders->allborders->border_style = zcl_excel_style_border=>c_border_thin.
    _mv_style_cell = lo_cell_style->get_guid( ).

  ENDMETHOD.

  METHOD _get_sorting_column.
    LOOP AT it_info ASSIGNING FIELD-SYMBOL(<lfs_info>).
      CHECK <lfs_info>-merge IS NOT INITIAL.
      IF ev_colpos1 IS INITIAL.
        ev_colpos1 = <lfs_info>-col_pos.
        CONTINUE.
      ENDIF.
      IF ev_colpos2 IS INITIAL.
        ev_colpos2 = <lfs_info>-col_pos.
        CONTINUE.
      ENDIF.
      IF ev_colpos3 IS INITIAL.
        ev_colpos3 = <lfs_info>-col_pos.
        CONTINUE.
      ENDIF.
      IF ev_colpos4 IS INITIAL.
        ev_colpos4 = <lfs_info>-col_pos.
        CONTINUE.
      ENDIF.
      IF ev_colpos5 IS INITIAL.
        ev_colpos5 = <lfs_info>-col_pos.
        CONTINUE.
      ENDIF.
      IF ev_colpos6 IS INITIAL.
        ev_colpos6 = <lfs_info>-col_pos.
        CONTINUE.
      ENDIF.
      IF ev_colpos7 IS INITIAL.
        ev_colpos7 = <lfs_info>-col_pos.
        CONTINUE.
      ENDIF.
      IF ev_colpos8 IS INITIAL.
        ev_colpos8 = <lfs_info>-col_pos.
        CONTINUE.
      ENDIF.
    ENDLOOP.

  ENDMETHOD.

  METHOD _data_to_excel.
    "data define
    """"""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""
    DATA :
       lo_worksheet TYPE REF TO zcl_excel_worksheet.

    DATA :
      lv_sub1beg TYPE i,
      lv_sub2beg TYPE i,
      lv_sub3beg TYPE i,
      lv_sub4beg TYPE i,
      lv_sub5beg TYPE i,
      lv_sub6beg TYPE i,
      lv_sub7beg TYPE i,
      lv_sub8beg TYPE i,
      lv_totbeg  TYPE i.

    DATA :
      lv_sub_count TYPE i VALUE 0.

    DATA :
      lv_row_from TYPE i,
      lv_row_to   TYPE i,
      lv_rowno    TYPE zexcel_cell_row,
      lv_style    TYPE zexcel_cell_style.

    "create worksheet
    """"""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""
    TRY.
        lo_worksheet = io_excel->get_active_worksheet( ).
        lo_worksheet->set_title( ip_title = CONV #( sy-title ) ).
      CATCH zcx_excel.
        EXIT.
    ENDTRY.

    "create header
    """"""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""
    LOOP AT it_info ASSIGNING FIELD-SYMBOL(<lfs_info>).
      TRY.
          lo_worksheet->set_cell( ip_row = 1
                                  ip_column = <lfs_info>-col_pos
                                  ip_value = <lfs_info>-text_m
                                  ip_style = _mv_style_title   ).
        CATCH zcx_excel.
      ENDTRY.
    ENDLOOP.

    "get sorting column number
    """"""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""
    me->_get_sorting_column(
      EXPORTING
        it_info    = it_info
      IMPORTING
        ev_colpos1 = DATA(lv_colpos1)
        ev_colpos2 = DATA(lv_colpos2)
        ev_colpos3 = DATA(lv_colpos3)
        ev_colpos4 = DATA(lv_colpos4)
        ev_colpos5 = DATA(lv_colpos5)
        ev_colpos6 = DATA(lv_colpos6)
        ev_colpos7 = DATA(lv_colpos7)
        ev_colpos8 = DATA(lv_colpos8)
    ).

    "set group begin line
    """"""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""
    lv_sub1beg = lv_sub2beg = lv_sub3beg = lv_sub4beg = lv_sub5beg = lv_sub6beg
    = lv_sub7beg = lv_sub8beg = lv_totbeg  = _mv_sub_begline.

    "get last row (use progress indicator )
    """"""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""
    DATA(lv_lines) = lines( it_data ).
    READ TABLE it_data ASSIGNING FIELD-SYMBOL(<lfs_data2>) INDEX lv_lines.
    DATA(lv_last_row) = <lfs_data2>-row_pos.

    "create body
    """"""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""
    LOOP AT it_data ASSIGNING FIELD-SYMBOL(<lgfs_row>) WHERE col_pos GT 0 GROUP BY <lgfs_row>-row_pos .

      LOOP AT GROUP <lgfs_row> ASSIGNING FIELD-SYMBOL(<lfs_column>) WHERE col_pos GT 0.

        "style
        """"""""""""""""""""""""""""""""""""""""""""""""""""""""""""""
        CASE <lfs_column>-style.
          WHEN c_red_cell.
            lv_style = _mv_style_red_cell.
          WHEN c_green_cell.
            lv_style = _mv_style_green_cell.
          WHEN c_subtotal.

            READ TABLE it_info ASSIGNING <lfs_info> WITH KEY col_pos = <lfs_column>-col_pos BINARY SEARCH.


            IF <lfs_column>-value IS INITIAL OR
                <lfs_info>-merge NE 'X'.
              lv_style = _mv_style_subtotal_topbottom.
            ELSE.
              READ TABLE it_data ASSIGNING FIELD-SYMBOL(<lfs_data>)
                  WITH KEY row_pos = <lfs_column>-row_pos + 1
                           col_pos = <lfs_column>-col_pos BINARY SEARCH.
              IF <lfs_data>-value EQ <lfs_column>-value.
                lv_style = _mv_style_subtotal_sides.
              ELSE.
                lv_style = _mv_style_subtotal_down.
              ENDIF.
            ENDIF.

          WHEN c_total.
            lv_style = _mv_style_total.
          WHEN OTHERS.

            IF <lfs_column>-col_pos EQ lv_colpos1 OR
               <lfs_column>-col_pos EQ lv_colpos2 OR
               <lfs_column>-col_pos EQ lv_colpos3 OR
               <lfs_column>-col_pos EQ lv_colpos4 OR
               <lfs_column>-col_pos EQ lv_colpos5 OR
               <lfs_column>-col_pos EQ lv_colpos6 OR
               <lfs_column>-col_pos EQ lv_colpos7 OR
               <lfs_column>-col_pos EQ lv_colpos8.

              lv_style = _mv_style_subtotal_sides.

            ELSE.

              lv_style = _mv_style_cell.

            ENDIF.
        ENDCASE.
        "row no
        """"""""""""""""""""""""""""""""""""""""""""""""""""""""""""""
        lv_rowno = <lgfs_row>-row_pos + c_header_count.

        "value
        """"""""""""""""""""""""""""""""""""""""""""""""""""""""""""""
        CONDENSE <lfs_column>-value NO-GAPS.

        "set cell
        """"""""""""""""""""""""""""""""""""""""""""""""""""""""""""""
        TRY.
            lo_worksheet->set_cell( ip_row = lv_rowno
                                    ip_column = <lfs_column>-col_pos
                                    ip_value = <lfs_column>-value
                                    ip_abap_type = cl_abap_typedescr=>typekind_string
                                    ip_style = COND #( WHEN lv_style IS NOT INITIAL THEN lv_style )  ).
          CATCH zcx_excel.
        ENDTRY.

        CLEAR : lv_style.
      ENDLOOP.

      "grouping
      """"""""""""""""""""""""""""""""""""""""""""""""""""""""""""""
      IF <lgfs_row>-style EQ c_subtotal.
        lv_sub_count = lv_sub_count + 1.

        CASE lv_sub_count.
          WHEN 1.
            lv_row_from = lv_sub1beg.
          WHEN 2.
            lv_row_from = lv_sub2beg.
          WHEN 3.
            lv_row_from = lv_sub3beg.
          WHEN 4.
            lv_row_from = lv_sub4beg.
          WHEN 5.
            lv_row_from = lv_sub5beg.
          WHEN 6.
            lv_row_from = lv_sub6beg.
          WHEN 7.
            lv_row_from = lv_sub7beg.
          WHEN 8.
            lv_row_from = lv_sub8beg.
          WHEN 9.
            lv_row_from = lv_totbeg.
        ENDCASE.

        lv_row_to = <lgfs_row>-row_pos .

        TRY.
            lo_worksheet->set_row_outline(
              EXPORTING
                iv_row_from  = lv_row_from
                iv_row_to    = lv_row_to
                iv_collapsed = abap_false
            ).
          CATCH zcx_excel.
        ENDTRY.

      ELSE.
        IF <lgfs_row>-style NE c_total.
          IF lv_sub_count NE 0.

            CASE lv_sub_count.
              WHEN 1.
                lv_sub1beg = <lgfs_row>-row_pos + c_header_count.
              WHEN 2.
                lv_sub1beg = <lgfs_row>-row_pos + c_header_count.
                lv_sub2beg = <lgfs_row>-row_pos + c_header_count.
              WHEN 3.
                lv_sub1beg = <lgfs_row>-row_pos + c_header_count.
                lv_sub2beg = <lgfs_row>-row_pos + c_header_count.
                lv_sub3beg = <lgfs_row>-row_pos + c_header_count.
              WHEN 4.
                lv_sub1beg = <lgfs_row>-row_pos + c_header_count.
                lv_sub2beg = <lgfs_row>-row_pos + c_header_count.
                lv_sub3beg = <lgfs_row>-row_pos + c_header_count.
                lv_sub4beg = <lgfs_row>-row_pos + c_header_count.
              WHEN 5.
                lv_sub1beg = <lgfs_row>-row_pos + c_header_count.
                lv_sub2beg = <lgfs_row>-row_pos + c_header_count.
                lv_sub3beg = <lgfs_row>-row_pos + c_header_count.
                lv_sub4beg = <lgfs_row>-row_pos + c_header_count.
                lv_sub5beg = <lgfs_row>-row_pos + c_header_count.
              WHEN 6.
                lv_sub1beg = <lgfs_row>-row_pos + c_header_count.
                lv_sub2beg = <lgfs_row>-row_pos + c_header_count.
                lv_sub3beg = <lgfs_row>-row_pos + c_header_count.
                lv_sub4beg = <lgfs_row>-row_pos + c_header_count.
                lv_sub5beg = <lgfs_row>-row_pos + c_header_count.
                lv_sub6beg = <lgfs_row>-row_pos + c_header_count.
              WHEN 7.
                lv_sub1beg = <lgfs_row>-row_pos + c_header_count.
                lv_sub2beg = <lgfs_row>-row_pos + c_header_count.
                lv_sub3beg = <lgfs_row>-row_pos + c_header_count.
                lv_sub4beg = <lgfs_row>-row_pos + c_header_count.
                lv_sub5beg = <lgfs_row>-row_pos + c_header_count.
                lv_sub6beg = <lgfs_row>-row_pos + c_header_count.
                lv_sub7beg = <lgfs_row>-row_pos + c_header_count.
              WHEN 8.
                lv_sub1beg = <lgfs_row>-row_pos + c_header_count.
                lv_sub2beg = <lgfs_row>-row_pos + c_header_count.
                lv_sub3beg = <lgfs_row>-row_pos + c_header_count.
                lv_sub4beg = <lgfs_row>-row_pos + c_header_count.
                lv_sub5beg = <lgfs_row>-row_pos + c_header_count.
                lv_sub6beg = <lgfs_row>-row_pos + c_header_count.
                lv_sub7beg = <lgfs_row>-row_pos + c_header_count.
                lv_sub8beg = <lgfs_row>-row_pos + c_header_count.
              WHEN 9 .
                lv_sub1beg = <lgfs_row>-row_pos + c_header_count.
                lv_sub2beg = <lgfs_row>-row_pos + c_header_count.
                lv_sub3beg = <lgfs_row>-row_pos + c_header_count.
                lv_sub4beg = <lgfs_row>-row_pos + c_header_count.
                lv_sub5beg = <lgfs_row>-row_pos + c_header_count.
                lv_sub6beg = <lgfs_row>-row_pos + c_header_count.
                lv_sub7beg = <lgfs_row>-row_pos + c_header_count.
                lv_sub8beg = <lgfs_row>-row_pos + c_header_count.
                lv_totbeg = <lgfs_row>-row_pos + c_header_count.
            ENDCASE.
          ENDIF.
          lv_sub_count = 0.
        ENDIF.
      ENDIF.
      CLEAR : lv_row_from,lv_row_to.

      _progress_indicator(
        EXPORTING
          iv_text       = TEXT-002
          iv_percentage = CONV i( <lgfs_row>-row_pos * 100 / lv_last_row )
      ).

    ENDLOOP.
  ENDMETHOD.
ENDCLASS.
