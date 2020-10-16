using Atomus.Diagnostics;
using DevExpress.XtraGrid;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Text;
using System.Windows.Forms;
using DevExpress.XtraGrid.Views.Base;
using DevExpress.Utils;
using DevExpress.XtraGrid.Views.Grid;
using DevExpress.XtraGrid.Columns;
using DevExpress.XtraGrid.Views.Grid.ViewInfo;
using DevExpress.XtraPrinting;
using DevExpress.Export;
using DevExpress.XtraEditors;
using DevExpress.XtraEditors.Repository;
using System.Data;
using System.Linq;

namespace Atomus.Control.Grid
{
    public class DevExpressGridControlAgent : IDataGridAgent
    {
        string[] menuListText;
        System.Windows.Forms.Control IDataGridAgent.GridControl { get; set; }
        Dictionary<TextEdit, FilterAttribute> FilterControl { get; set; }
        int headerRowCount;
        int headerHeight;

        EditAble IDataGridAgent.EditAble
        {
            get
            {
                return ((GridView)((GridControl)((IDataGridAgent)this).GridControl).MainView).OptionsBehavior.Editable ? EditAble.True : EditAble.False;
            }
            set
            {
                ((GridView)((GridControl)((IDataGridAgent)this).GridControl).MainView).OptionsBehavior.Editable = (value == EditAble.True);
                ((GridView)((GridControl)((IDataGridAgent)this).GridControl).MainView).OptionsSelection.EnableAppearanceFocusedCell = (value == EditAble.True);
            }
        }
        AddRows IDataGridAgent.AddRows
        {
            get
            {
                return ((GridView)((GridControl)((IDataGridAgent)this).GridControl).MainView).OptionsBehavior.AllowAddRows == DefaultBoolean.True ? AddRows.True : AddRows.False;
            }
            set
            {
                if (value == AddRows.True)
                    ((GridView)((GridControl)((IDataGridAgent)this).GridControl).MainView).OptionsBehavior.AllowAddRows = DefaultBoolean.True;
                else
                    ((GridView)((GridControl)((IDataGridAgent)this).GridControl).MainView).OptionsBehavior.AllowAddRows = DefaultBoolean.False;
            }
        }
        DeleteRows IDataGridAgent.DeleteRows
        {
            get
            {
                return ((GridView)((GridControl)((IDataGridAgent)this).GridControl).MainView).OptionsBehavior.AllowDeleteRows == DefaultBoolean.True ? DeleteRows.True : DeleteRows.False;
            }
            set
            {
                if (value == DeleteRows.True)
                    ((GridView)((GridControl)((IDataGridAgent)this).GridControl).MainView).OptionsBehavior.AllowDeleteRows = DefaultBoolean.True;
                else
                    ((GridView)((GridControl)((IDataGridAgent)this).GridControl).MainView).OptionsBehavior.AllowDeleteRows = DefaultBoolean.False;
            }
        }
        ResizeRows IDataGridAgent.ResizeRows
        {
            get
            {
                return ((GridView)((GridControl)((IDataGridAgent)this).GridControl).MainView).OptionsCustomization.AllowRowSizing ? ResizeRows.True : ResizeRows.False;
            }
            set
            {
                ((GridView)((GridControl)((IDataGridAgent)this).GridControl).MainView).OptionsCustomization.AllowRowSizing = (value == ResizeRows.True) ? true : false;
            }
        }
        AutoSizeColumns IDataGridAgent.AutoSizeColumns
        {
            get
            {
                return ((GridView)((GridControl)((IDataGridAgent)this).GridControl).MainView).OptionsView.ColumnAutoWidth ? AutoSizeColumns.True : AutoSizeColumns.False;
            }
            set
            {
                ((GridView)((GridControl)((IDataGridAgent)this).GridControl).MainView).OptionsView.ColumnAutoWidth = (value == AutoSizeColumns.True) ? true : false;
            }
        }
        AutoSizeRows IDataGridAgent.AutoSizeRows
        {
            get
            {
                return ((GridView)((GridControl)((IDataGridAgent)this).GridControl).MainView).OptionsView.RowAutoHeight ? AutoSizeRows.True : AutoSizeRows.False;
            }
            set
            {
                ((GridView)((GridControl)((IDataGridAgent)this).GridControl).MainView).OptionsView.RowAutoHeight = (value == AutoSizeRows.True) ? true : false;
            }
        }
        ColumnsHeadersVisible IDataGridAgent.ColumnsHeadersVisible
        {
            get
            {
                return ((GridView)((GridControl)((IDataGridAgent)this).GridControl).MainView).OptionsView.ShowColumnHeaders ? ColumnsHeadersVisible.True : ColumnsHeadersVisible.False;
            }
            set
            {
                ((GridView)((GridControl)((IDataGridAgent)this).GridControl).MainView).OptionsView.ShowColumnHeaders = (value == ColumnsHeadersVisible.True) ? true : false;
            }
        }
        EnableMenu IDataGridAgent.EnableMenu
        {
            get
            {
                if (((IDataGridAgent)this).GridControl.ContextMenu != null)
                    return EnableMenu.False;
                else
                    return EnableMenu.False;
            }
            set
            {
                if (value == EnableMenu.True && ((GridView)((GridControl)((IDataGridAgent)this).GridControl).MainView).GridControl.ContextMenu == null)
                    SetContextMenuStrip((GridView)((GridControl)((IDataGridAgent)this).GridControl).MainView, this.menuListText);
            }
        }
        MultiSelect IDataGridAgent.MultiSelect
        {
            get
            {
                return ((GridView)((GridControl)((IDataGridAgent)this).GridControl).MainView).OptionsSelection.MultiSelect ? MultiSelect.True : MultiSelect.False;
            }
            set
            {
                ((GridView)((GridControl)((IDataGridAgent)this).GridControl).MainView).OptionsSelection.MultiSelect = (value == MultiSelect.True) ? true : false;
            }
        }
        Alignment IDataGridAgent.HeadAlignment
        {
            get
            {
                switch (((GridView)((GridControl)((IDataGridAgent)this).GridControl).MainView).Appearance.HeaderPanel.TextOptions.HAlignment)
                {
                    case HorzAlignment.Center:
                        return Alignment.MiddleCenter;
                    case HorzAlignment.Default:
                        return Alignment.NotSet;
                    case HorzAlignment.Far:
                        return Alignment.MiddleRight;
                    case HorzAlignment.Near:
                        return Alignment.MiddleLeft;
                    default:
                        throw new AtomusException("Alignment Not Support.");
                }
            }
            set
            {
                switch (value)
                {
                    case Alignment.MiddleCenter:
                        ((GridView)((GridControl)((IDataGridAgent)this).GridControl).MainView).Appearance.HeaderPanel.TextOptions.HAlignment = HorzAlignment.Center;
                        break;
                    case Alignment.NotSet:
                        ((GridView)((GridControl)((IDataGridAgent)this).GridControl).MainView).Appearance.HeaderPanel.TextOptions.HAlignment = HorzAlignment.Default;
                        break;
                    case Alignment.MiddleRight:
                        ((GridView)((GridControl)((IDataGridAgent)this).GridControl).MainView).Appearance.HeaderPanel.TextOptions.HAlignment = HorzAlignment.Far;
                        break;
                    case Alignment.MiddleLeft:
                        ((GridView)((GridControl)((IDataGridAgent)this).GridControl).MainView).Appearance.HeaderPanel.TextOptions.HAlignment = HorzAlignment.Near;
                        break;
                    default:
                        throw new AtomusException("Alignment Not Support.");
                }
            }
        }
        int IDataGridAgent.HeaderHeight
        {
            get
            {
                return this.headerHeight;
            }
            set
            {
                this.headerHeight = value;
                this.SetHeaderHeight();
            }
        }
        int IDataGridAgent.HeaderRowCount
        {
            get
            {
                return this.headerRowCount;
            }
            set
            {
                this.headerRowCount = value;
                this.SetHeaderHeight();
            }
        }
        private void SetHeaderHeight()
        {
            if (((IDataGridAgent)this).HeaderRowCount > 0)
                if (((IDataGridAgent)this).HeaderHeight > 0)//'행 높이
                    ((GridView)((GridControl)((IDataGridAgent)this).GridControl).MainView).ColumnPanelRowHeight = ((IDataGridAgent)this).HeaderRowCount * ((IDataGridAgent)this).HeaderHeight;
                else
                    ((GridView)((GridControl)((IDataGridAgent)this).GridControl).MainView).ColumnPanelRowHeight = ((IDataGridAgent)this).HeaderRowCount * ((IDataGridAgent)this).RowHeight;
        }

        int IDataGridAgent.RowHeight
        {
            get
            {
                return ((GridView)((GridControl)((IDataGridAgent)this).GridControl).MainView).RowHeight;
            }
            set
            {
                if (value >= 0)//'행 높이
                    ((GridView)((GridControl)((IDataGridAgent)this).GridControl).MainView).RowHeight = value;
            }
        }
        RowHeadersVisible IDataGridAgent.RowHeadersVisible
        {
            get
            {
                return ((GridView)((GridControl)((IDataGridAgent)this).GridControl).MainView).OptionsView.ShowIndicator ? RowHeadersVisible.True : RowHeadersVisible.False;
            }
            set
            {
                ((GridView)((GridControl)((IDataGridAgent)this).GridControl).MainView).OptionsView.ShowIndicator = (value == RowHeadersVisible.True) ? true : false;
            }
        }
        Selection IDataGridAgent.Selection
        {
            get
            {
                switch (((GridView)((GridControl)((IDataGridAgent)this).GridControl).MainView).OptionsSelection.MultiSelectMode)
                {
                    case GridMultiSelectMode.CellSelect:
                        return Selection.CellSelect;
                    case GridMultiSelectMode.RowSelect:
                        return Selection.FullRowSelect;
                    case GridMultiSelectMode.CheckBoxRowSelect:
                        return Selection.RowHeaderSelect;
                    default:
                        throw new AtomusException("Selection Not Support.");
                }
            }
            set
            {
                switch (value)
                {
                    case Selection.CellSelect:
                        ((GridView)((GridControl)((IDataGridAgent)this).GridControl).MainView).OptionsSelection.MultiSelectMode = GridMultiSelectMode.CellSelect;
                        break;
                    case Selection.FullRowSelect:
                        ((GridView)((GridControl)((IDataGridAgent)this).GridControl).MainView).OptionsSelection.MultiSelectMode = GridMultiSelectMode.RowSelect;
                        break;
                    case Selection.RowHeaderSelect:
                        ((GridView)((GridControl)((IDataGridAgent)this).GridControl).MainView).OptionsSelection.MultiSelectMode = GridMultiSelectMode.CheckBoxRowSelect;
                        break;
                    default:
                        throw new AtomusException("Alignment Not Support.");
                }
            }
        }

        public DevExpressGridControlAgent()
        {
            this.FilterControl = new Dictionary<TextEdit, FilterAttribute>();

            try
            {
                this.menuListText = this.GetAttribute("MenuListText").Split(',').Translate();
            }
            catch (Exception exception)
            {
                DiagnosticsTool.MyTrace(exception);
            }
        }

        void IDataGridAgent.Init(EditAble editAble, AddRows allowAddRows, DeleteRows allowDeleteRows, ResizeRows allowResizeRows, AutoSizeColumns autoSizeColumns
            , AutoSizeRows autoSizeRows, ColumnsHeadersVisible columnsHeadersVisible, EnableMenu enableMenu, MultiSelect multiSelect, Alignment headAlign, int headerHeight, int headerRowCount
            , int rowHeight, RowHeadersVisible rowHeadersVisible, Selection selection)
        {
            GridView gridView;

            gridView = ((GridView)((GridControl)((IDataGridAgent)this).GridControl).MainView);

            gridView.OptionsView.RowAutoHeight = true; //자동 줄바꿈 금지
            gridView.OptionsView.ColumnHeaderAutoHeight = DefaultBoolean.True;//컬럼헤더 높이 변경 금지
            ((IDataGridAgent)this).EditAble = editAble;
            ((IDataGridAgent)this).AddRows = allowAddRows;
            ((IDataGridAgent)this).DeleteRows = allowDeleteRows;
            ((IDataGridAgent)this).ResizeRows = allowResizeRows;
            ((IDataGridAgent)this).AutoSizeColumns = autoSizeColumns;
            ((IDataGridAgent)this).AutoSizeRows = autoSizeRows;
            ((IDataGridAgent)this).ColumnsHeadersVisible = columnsHeadersVisible;
            ((IDataGridAgent)this).MultiSelect = multiSelect;
            ((IDataGridAgent)this).EnableMenu = enableMenu;
            ((IDataGridAgent)this).HeadAlignment = headAlign;
            ((IDataGridAgent)this).RowHeight = rowHeight;

            ((IDataGridAgent)this).HeaderRowCount = headerRowCount;//'컬럼헤더 수
            ((IDataGridAgent)this).HeaderHeight = headerHeight;

            ((IDataGridAgent)this).RowHeadersVisible = rowHeadersVisible;
            ((IDataGridAgent)this).Selection = selection;

            gridView.GridControl.DoubleBuffered(true);

            gridView.DataSourceChanged -= this.GridView_DataSourceChanged;
            gridView.DataSourceChanged += this.GridView_DataSourceChanged;

            gridView.RowCountChanged -= this.DataGridView_DataBindingComplete;
            gridView.RowCountChanged += this.DataGridView_DataBindingComplete;

            this.SetSkin(gridView);
        }

        void IDataGridAgent.AddColumn(int width, ColumnVisible visible, EditAble editAble, Filter allowFilter, Merge allowMerge, Sort sortMode, object editControl, Alignment textAlign, string format, string name, params string[] caption)
        {
            GridView gridView;
            GridColumn gridColumn;

            //타입이 BandedGridView 라면
            if (((GridControl)((IDataGridAgent)this).GridControl).MainView.GetType().Equals(typeof(DevExpress.XtraGrid.Views.BandedGrid.BandedGridView)))
            {
                DevExpress.XtraGrid.Views.BandedGrid.GridBand gridBand;
                DevExpress.XtraGrid.Views.BandedGrid.BandedGridView bandedGridView;
                DevExpress.XtraGrid.Views.BandedGrid.BandedGridColumn bandedGridColumn; 

                //MainView
                bandedGridView = ((DevExpress.XtraGrid.Views.BandedGrid.BandedGridView)((GridControl)((IDataGridAgent)this).GridControl).MainView);

                caption = caption.Translate();

                //현재 칼럼의 가장 대분류 캡션 가져오기
                gridBand = bandedGridView.Bands[caption[0]];


                //그리드에 해당 밴드가 없다면 생성
                if (gridBand == null)
                {
                    //밴드 생성
                    gridBand = new DevExpress.XtraGrid.Views.BandedGrid.GridBand();
                    
                    //그리드뷰 밴드 추가
                    bandedGridView.Bands.Add(gridBand);
                }

                //기본값 할당
                gridBand.Caption = caption[0];
                gridBand.Name = caption[0];
                //gridBand.AutoFillDown = true;
                //gridBand.RowCount = 1;
                bandedGridView.Appearance.BandPanel.TextOptions.HAlignment = HorzAlignment.Center;
                
                //대분류 밴드 참조
                DevExpress.XtraGrid.Views.BandedGrid.GridBand gridBand_temp = gridBand;

                //가져온 밴드에 속성 할당(합병 작업)
                for (int idx = 0; idx <= caption.Length - 2 && idx <= ((IDataGridAgent)this).HeaderRowCount - 2; idx++)
                {
                    //다음 캡션이 같은 값이라 합병할때
                    if (caption[idx] == caption[idx + 1])
                    {
                        gridBand_temp.Caption = caption[idx];
                        gridBand_temp.Name = caption[idx];

                        gridBand_temp.AutoFillDown = true;
                        gridBand_temp.RowCount = idx + 1;
                        //gridBand_temp.RowCount += 1;

                        gridBand_temp.Visible = (visible == ColumnVisible.True);
                    }
                    else // 다음 캡션이 다른 값일 때
                    {
                        DevExpress.XtraGrid.Views.BandedGrid.GridBand gridBand_Child = gridBand_temp.Children[caption[idx + 1]];
                        
                        if (gridBand_Child == null)
                        {
                            gridBand_Child = bandedGridView.Bands.Add();
                            gridBand_temp.Children.Add(gridBand_Child);
                        }
                        gridBand_Child.Visible = (visible == ColumnVisible.True);

                        gridBand_Child.AutoFillDown = false;
                        gridBand_Child.RowCount = 1;
                        gridBand_Child.Caption = caption[idx + 1];
                        gridBand_Child.Name = caption[idx + 1];

                        //temp bnad 값 참조
                        gridBand_temp = gridBand_Child;
                    }
                }

                //밴드에는 데이터 바인딩이 되지 않으므로
                //마지막 밴드 하위에 columnheader visible false 두고 컬럼값 할당
                bandedGridColumn = gridBand_temp.Columns[gridBand_temp.Caption];
                if(bandedGridColumn == null)
                {
                    bandedGridColumn = bandedGridView.Columns.Add();
                }
                gridBand_temp.Columns.Add(bandedGridColumn);

                bandedGridColumn.FieldName = name;
                bandedGridColumn.Name = name; ;
                bandedGridColumn.Caption = gridBand_temp.Caption;
                bandedGridColumn.Visible = (visible == ColumnVisible.True);
                bandedGridColumn.Width = width;
                bandedGridColumn.OptionsColumn.AllowEdit = (editAble == EditAble.True);
                bandedGridColumn.OptionsFilter.AllowFilter = (allowFilter == Filter.True);
                bandedGridColumn.OptionsColumn.AllowMerge = (allowMerge == Merge.True ? (DefaultBoolean)0 : (DefaultBoolean)1);
                if(bandedGridColumn.OptionsColumn.AllowMerge == DefaultBoolean.True)
                {
                    bandedGridView.OptionsView.AllowCellMerge = true;
                }
                bandedGridView.OptionsView.ShowColumnHeaders = false;
                bandedGridColumn.OptionsColumn.AllowShowHide = true;


                if (editControl != null)
                {
                    ((GridControl)((IDataGridAgent)this).GridControl).RepositoryItems.Add((RepositoryItem)editControl);

                    bandedGridColumn.ColumnEdit = (RepositoryItem)editControl;

                    if (editControl is RepositoryItemCheckEdit)
                    {
                        ((RepositoryItemCheckEdit)editControl).ValueChecked = 1;
                        ((RepositoryItemCheckEdit)editControl).ValueUnchecked = 0;

                        bandedGridView.OptionsSelection.ShowCheckBoxSelectorInColumnHeader = (caption == null || caption.Count() < 1 || caption[0].Trim().Length < 1) ? DefaultBoolean.True : DefaultBoolean.False;
                        bandedGridView.OptionsSelection.CheckBoxSelectorField = name;

                        bandedGridView.OptionsSelection.MultiSelectMode = GridMultiSelectMode.CheckBoxRowSelect;
                    }

                    //ShowColumnHeaderCheckBox(GridControl, ColumnName, False)

                    //If TypeOf(RepositoryItem) Is XtraEditors.Repository.RepositoryItemCheckEdit Then
                    //    _RepositoryItem = DirectCast(RepositoryItem, RepositoryItemCheckEdit)
                    //    _RepositoryItem.ValueChecked = 1
                    //    _RepositoryItem.ValueUnchecked = 0
                    //    .DisplayFormat.FormatType = FormatType.Numeric
                    //    _dataColumn.DataType = GetType(Integer)

                    //    AddHandler _grdView.CustomDrawColumnHeader, AddressOf CustomDrawColumnHeader
                    //    AddHandler _grdView.Click, AddressOf ViewClick
                    //End If
                }



                if (Sort.Automatic == sortMode)
                    bandedGridColumn.SortMode = ColumnSortMode.Default;
                else
                    if (Sort.NotSortable == sortMode)
                    bandedGridColumn.OptionsColumn.AllowSort = DefaultBoolean.False;
                else
                    if (Sort.Programmatic == sortMode)
                    bandedGridColumn.SortMode = ColumnSortMode.Custom;



                switch (textAlign)
                {
                    case Alignment.MiddleLeft:
                        bandedGridColumn.AppearanceCell.TextOptions.HAlignment = HorzAlignment.Near;
                        break;
                    case Alignment.MiddleCenter:
                        bandedGridColumn.AppearanceCell.TextOptions.HAlignment = HorzAlignment.Center;
                        break;
                    case Alignment.MiddleRight:
                        bandedGridColumn.AppearanceCell.TextOptions.HAlignment = HorzAlignment.Far;
                        break;
                    default:
                        bandedGridColumn.AppearanceCell.TextOptions.HAlignment = HorzAlignment.Default;
                        break;
                }

                bandedGridColumn.AppearanceCell.Options.UseTextOptions = true;

                if (!format.Equals(string.Empty))
                    bandedGridColumn.DisplayFormat.FormatType = FormatType.Custom;

                bandedGridColumn.DisplayFormat.FormatString = format;


            }
            else
            {
                gridView = ((GridView)((GridControl)((IDataGridAgent)this).GridControl).MainView);

                if (gridView.Columns[name] != null)//'기존에 컬럼이 있으면 가져오기
                    gridColumn = gridView.Columns[name];
                else //'교체할 컬럼이 없으면 생성
                {
                    gridColumn = new GridColumn();
                    gridView.Columns.Add(gridColumn);
                }

                gridColumn.FieldName = name;//'Data 컬럼명
                gridColumn.Name = name;

                caption = caption.Translate();

                if (((IDataGridAgent)this).HeaderRowCount == 1)//'컬럼 헤더 수가 1개면
                    gridColumn.Caption = caption[caption.Length - 1];
                else
                {
                    gridColumn.Caption = "";

                    for (int i = 0; i <= ((IDataGridAgent)this).HeaderRowCount - 1; i++)
                        if (i < caption.Length)
                            gridColumn.Caption += caption[i] + Environment.NewLine;
                        else
                            gridColumn.Caption += Environment.NewLine;

                    if (gridColumn.Caption.Length > 0)
                        gridColumn.Caption = gridColumn.Caption.Substring(0, gridColumn.Caption.Length - 2);// '마지막 vbCrLf 제거
                }

                gridColumn.Visible = (visible == ColumnVisible.True);
                gridColumn.Width = width;
                gridColumn.OptionsColumn.AllowEdit = (editAble == EditAble.True);
                gridColumn.OptionsFilter.AllowFilter = (allowFilter == Filter.True);
                gridColumn.OptionsColumn.AllowMerge = (allowMerge == Merge.True ? (DefaultBoolean)0 : (DefaultBoolean)1);
                if (gridColumn.OptionsColumn.AllowMerge == DefaultBoolean.True)
                {
                    gridView.OptionsView.AllowCellMerge = true;
                }
                //gridView.OptionsView.ShowColumnHeaders = false;
                //gridColumn.OptionsColumn.AllowShowHide = true;

                if (editControl != null)
                {
                    ((GridControl)((IDataGridAgent)this).GridControl).RepositoryItems.Add((RepositoryItem)editControl);

                    gridColumn.ColumnEdit = (RepositoryItem)editControl;

                    if (editControl is RepositoryItemCheckEdit)
                    {
                        ((RepositoryItemCheckEdit)editControl).ValueChecked = 1;
                        ((RepositoryItemCheckEdit)editControl).ValueUnchecked = 0;

                        gridView.OptionsSelection.ShowCheckBoxSelectorInColumnHeader = (caption == null || caption.Count() < 1 || caption[0].Trim().Length < 1)? DefaultBoolean.True : DefaultBoolean.False;
                        gridView.OptionsSelection.CheckBoxSelectorField = name;

                        gridView.OptionsSelection.MultiSelectMode = GridMultiSelectMode.CheckBoxRowSelect;
                    }

                    //ShowColumnHeaderCheckBox(GridControl, ColumnName, False)

                    //If TypeOf(RepositoryItem) Is XtraEditors.Repository.RepositoryItemCheckEdit Then
                    //    _RepositoryItem = DirectCast(RepositoryItem, RepositoryItemCheckEdit)
                    //    _RepositoryItem.ValueChecked = 1
                    //    _RepositoryItem.ValueUnchecked = 0
                    //    .DisplayFormat.FormatType = FormatType.Numeric
                    //    _dataColumn.DataType = GetType(Integer)

                    //    AddHandler _grdView.CustomDrawColumnHeader, AddressOf CustomDrawColumnHeader
                    //    AddHandler _grdView.Click, AddressOf ViewClick
                    //End If
                }

                if (Sort.Automatic == sortMode)
                    gridColumn.SortMode = ColumnSortMode.Default;
                else
                    if (Sort.NotSortable == sortMode)
                    gridColumn.OptionsColumn.AllowSort = DefaultBoolean.False;
                else
                    if (Sort.Programmatic == sortMode)
                    gridColumn.SortMode = ColumnSortMode.Custom;

                //if (gridView.Columns[name] == null)
                //    gridView.Columns.Add(gridColumn);

                switch (textAlign)
                {
                    case Alignment.MiddleLeft:
                        gridColumn.AppearanceCell.TextOptions.HAlignment = HorzAlignment.Near;
                        break;
                    case Alignment.MiddleCenter:
                        gridColumn.AppearanceCell.TextOptions.HAlignment = HorzAlignment.Center;
                        break;
                    case Alignment.MiddleRight:
                        gridColumn.AppearanceCell.TextOptions.HAlignment = HorzAlignment.Far;
                        break;
                    default:
                        gridColumn.AppearanceCell.TextOptions.HAlignment = HorzAlignment.Default;
                        break;
                }

                gridColumn.AppearanceCell.Options.UseTextOptions = true;

                if (!format.Equals(string.Empty))
                    gridColumn.DisplayFormat.FormatType = FormatType.Custom;

                gridColumn.DisplayFormat.FormatString = format;
            }

            
        }

        void IDataGridAgent.Clear()
        {
            GridView gridView;
            DevExpress.XtraGrid.Views.BandedGrid.BandedGridView bandedGridView;

            //타입이 BandedGridView 라면
            if (((GridControl)((IDataGridAgent)this).GridControl).MainView.GetType().Equals(typeof(DevExpress.XtraGrid.Views.BandedGrid.BandedGridView)))
            {
                //MainView
                bandedGridView = ((DevExpress.XtraGrid.Views.BandedGrid.BandedGridView)((GridControl)((IDataGridAgent)this).GridControl).MainView);
                bandedGridView.Columns.Clear();
                bandedGridView.Bands.Clear();
            }
            else
            {
                gridView = ((GridView)((GridControl)((IDataGridAgent)this).GridControl).MainView);
                gridView.Columns.Clear();
                gridView.GridControl.DataSource = null;
            }
        }

        void IDataGridAgent.RemoveColumn(string name)
        {
            GridView gridView;
            GridColumn gridColumn;

            gridView = ((GridView)((GridControl)((IDataGridAgent)this).GridControl).MainView);

            gridColumn = gridView.Columns[name];

            gridView.Columns.Remove(gridColumn);
        }

        void IDataGridAgent.RemoveColumn(int index)
        {
            GridView gridView;

            gridView = ((GridView)((GridControl)((IDataGridAgent)this).GridControl).MainView);

            gridView.Columns.Remove(gridView.Columns[index]);
        }

        private void SetContextMenuStrip(GridView gridView, string[] menuListText)
        {
            ContextMenuStrip contextMenuStrip;
            ToolStripMenuItem toolStripMenuItem;

            contextMenuStrip = new ContextMenuStrip();

            contextMenuStrip.Opening += ContextMenuStrip_Opening;

            toolStripMenuItem = null;

            for (int i = 0; i <= this.menuListText.Length - 1; i++)
            {
                if (this.menuListText[i] != "")
                {
                    toolStripMenuItem = new ToolStripMenuItem(this.menuListText[i], null, ToolStripMenuItem_Click);
                    contextMenuStrip.Items.Add(toolStripMenuItem);
                }
                else
                    contextMenuStrip.Items.Add(new ToolStripSeparator());

                switch (i)
                {
                    case 3://복사
                        if (gridView.OptionsSelection.MultiSelect == false)
                            toolStripMenuItem.Enabled = false;
                        break;

                    case 5://행추가
                        if (gridView.OptionsBehavior.AllowAddRows == DefaultBoolean.False)
                            toolStripMenuItem.Enabled = false;
                        break;

                    case 6://행삭제
                        if (gridView.OptionsBehavior.AllowDeleteRows == DefaultBoolean.False)
                            toolStripMenuItem.Enabled = false;
                        break;

                    case 7://행복사
                        if (gridView.OptionsBehavior.AllowAddRows == DefaultBoolean.False)
                            toolStripMenuItem.Enabled = false;
                        break;
                }
            }

            ((GridView)((GridControl)((IDataGridAgent)this).GridControl).MainView).GridControl.ContextMenuStrip = contextMenuStrip;
        }

        private void ContextMenuStrip_Opening(object sender, System.ComponentModel.CancelEventArgs e)
        {
            ContextMenuStrip contextMenuStrip;

            GridView gridView;
            decimal sum;
            decimal count;
            decimal avg;

            sum = 0;
            count = 0;
            avg = 0;

            try
            {
                gridView = gridView = ((GridView)((GridControl)((IDataGridAgent)this).GridControl).MainView);
                contextMenuStrip = (ContextMenuStrip)sender;

                if (gridView.GetSelectedCells().Length > 0)
                    foreach (GridCell gridCell in gridView.GetSelectedCells())
                    {
                        try
                        {
                            if (!gridCell.Column.Visible)
                                continue;

                            sum += gridView.GetDataRow(gridCell.RowHandle)[gridCell.Column.Name].ToString().ToDecimal();
                            count += 1;
                        }
                        catch (Exception exception)
                        {
                            DiagnosticsTool.MyTrace(exception);
                        }
                    }
                else  if (gridView.GetSelectedRows().Length > 0)
                    foreach (int handles in gridView.GetSelectedRows())
                    {
                        foreach (GridColumn gridColumn in gridView.Columns)
                        {
                            try
                            {
                                if (!gridColumn.Visible)
                                    continue;

                                sum += gridView.GetDataRow(handles)[gridColumn.Name].ToString().ToDecimal();
                                count += 1;
                            }
                            catch (Exception exception)
                            {
                                DiagnosticsTool.MyTrace(exception);
                            }
                        }
                    }

                avg = sum / count;

                contextMenuStrip.Items[9].Text = string.Format("Sum : {0}", sum);
                contextMenuStrip.Items[10].Text = string.Format("Avg : {0}", avg);

                contextMenuStrip.Items[9].Tag = sum;
                contextMenuStrip.Items[10].Tag = avg;
            }
            catch (Exception exception)
            {
                DiagnosticsTool.MyTrace(exception);

                contextMenuStrip = (ContextMenuStrip)sender;
                contextMenuStrip.Items[9].Text = string.Format("Sum : {0}", 0);
                contextMenuStrip.Items[10].Text = string.Format("Avg : {0}", 0);

                contextMenuStrip.Items[9].Tag = 0;
                contextMenuStrip.Items[10].Tag = 0;
            }
        }

        private void ToolStripMenuItem_Click(Object _sender, EventArgs e)
        {
            ContextMenuStrip contextMenuStrip;
            ToolStripMenuItem toolStripMenuItem;
            SaveFileDialog fileDialog;
            GridView gridView;
            Process process;

            gridView = gridView = ((GridView)((GridControl)((IDataGridAgent)this).GridControl).MainView);

            toolStripMenuItem = (ToolStripMenuItem)_sender;
            contextMenuStrip = (ContextMenuStrip)toolStripMenuItem.Owner;
            gridView.GridControl = ((GridView)((GridControl)((IDataGridAgent)this).GridControl).MainView).GridControl.ContextMenuStrip.SourceControl as GridControl;

            //엑셀 저장, 엑셀 저장 & 열기
            if (toolStripMenuItem.Text.Equals(this.menuListText[0]) || toolStripMenuItem.Text.Equals(this.menuListText[1]))
            {
                fileDialog = new SaveFileDialog()
                {
                    DefaultExt = "*.xlsx",
                    Filter = "xlsx Files(.xlsx)|*.xlsx|xls files (*.xls)|*.xls|All files (*.*)|*.*"
                };

                if (fileDialog.ShowDialog() == DialogResult.OK)
                {
                    XlsxExportOptionsEx op = new XlsxExportOptionsEx();

                    if (((GridControl)((IDataGridAgent)this).GridControl).MainView.GetType().Equals(typeof(DevExpress.XtraGrid.Views.BandedGrid.BandedGridView)))
                        op.ShowColumnHeaders = DefaultBoolean.False;//op.AllowCombinedBandAndColumnHeaderCellMerge = DefaultBoolean.True;

                    foreach (GridColumn gridColumn in gridView.Columns)
                        if (gridColumn.ColumnType == (new byte[0]).GetType())
                        {
                            op.ExportType = ExportType.WYSIWYG;
                            break;
                        }

                    gridView.GridControl.ExportToXlsx(fileDialog.FileName, op);

                    if (toolStripMenuItem.Text.Equals(this.menuListText[1])) //'엑셀 저장 & 열기
                    {
                        process = new Process();
                        process = Process.Start(fileDialog.FileName);
                    }
                }

                return;
            }

            //출력
            if (toolStripMenuItem.Text.Equals(menuListText[2]))
            {
                if (((GridControl)((IDataGridAgent)this).GridControl).MainView.GetType().Equals(typeof(DevExpress.XtraGrid.Views.BandedGrid.BandedGridView)))
                    gridView.OptionsPrint.PrintHeader = false;

                gridView.ShowPrintPreview();

                return;
            }


            //복사
            if (toolStripMenuItem.Text.Equals(menuListText[3]))
            {
                gridView.OptionsSelection.MultiSelect = true;
                gridView.SelectAll();
                gridView.CopyToClipboard();
                gridView.ClearSelection();

                return;
            }


            //행추가
            if (toolStripMenuItem.Text.Equals(menuListText[5]))
            {
                if (gridView.DataSource != null && gridView.DataSource is System.Data.DataView)
                {
                    if (gridView.OptionsBehavior.AllowAddRows == DefaultBoolean.True)
                    {
                        ((System.Data.DataView)gridView.DataSource).AddNew().EndEdit();
                    }
                }

                return;
            }


            //행삭제
            if (toolStripMenuItem.Text.Equals(menuListText[6]))
            {
                if (gridView.DataSource != null && gridView.DataSource is System.Data.DataView)
                    if (gridView.OptionsBehavior.AllowDeleteRows == DefaultBoolean.True)
                        ((System.Data.DataView)gridView.DataSource).Delete(gridView.GetFocusedDataSourceRowIndex());

                return;
            }

            //행복사
            if (toolStripMenuItem.Text.Equals(menuListText[7]))
            {
                gridView.GetSelectedCells();
                gridView.CopyToClipboard();

                return;
            }

            //Sum
            if (toolStripMenuItem.Equals(contextMenuStrip.Items[9]))
            {
                Clipboard.SetText(((decimal)toolStripMenuItem.Tag).ToString());
            }

            //Avg
            if (toolStripMenuItem.Equals(contextMenuStrip.Items[10]))
            {
                Clipboard.SetText(((decimal)toolStripMenuItem.Tag).ToString());
            }
        }
        void IDataGridAgent.AddColumnFiter(SearchAll searchAll, StartsWith startsWith, AutoComplete autoComplete, string name, params System.Windows.Forms.Control[] controls)
        {
            GridView gridView;
            FilterAttribute fiterAttribute;
            TextEdit textEdit;

            gridView = gridView = ((GridView)((GridControl)((IDataGridAgent)this).GridControl).MainView);

            fiterAttribute = new FilterAttribute()
            {
                GridView = gridView,
                ColumnName = name,
                IsSearchAll = (searchAll == SearchAll.True),
                IsStartsWith = (startsWith == StartsWith.True),
                AutoCompleteMode = (AutoCompleteMode)Enum.Parse(typeof(AutoCompleteMode), autoComplete.ToString()),
                AutoCompleteStringCollection = new AutoCompleteStringCollection()
            };

            foreach (System.Windows.Forms.Control control in controls)
            {
                if (control is TextEdit)
                {
                    textEdit = (TextEdit)control;

                    this.FilterControl.Add(textEdit, fiterAttribute);

                    if (fiterAttribute.AutoCompleteMode != AutoCompleteMode.None)
                    {
                        textEdit.TextChanged -= this.Control_TextChanged;
                        textEdit.TextChanged += this.Control_TextChanged;

                        textEdit.MaskBox.AutoCompleteMode = fiterAttribute.AutoCompleteMode;
                        textEdit.MaskBox.AutoCompleteCustomSource = fiterAttribute.AutoCompleteStringCollection;
                        textEdit.MaskBox.AutoCompleteSource = AutoCompleteSource.CustomSource;
                    }

                    textEdit.DoubleBuffered(true);
                }
            }
        }
        void IDataGridAgent.AddColumnFiter(SearchAll searchAll, StartsWith startsWith, AutoComplete autoComplete, int index, params System.Windows.Forms.Control[] controls)
        {
            GridView gridView;

            gridView = gridView = ((GridView)((GridControl)((IDataGridAgent)this).GridControl).MainView);

            ((IDataGridAgent)this).AddColumnFiter(searchAll, startsWith, autoComplete, gridView.Columns[index].Name, controls);
        }

        void IDataGridAgent.RemoveColumnFiter(params System.Windows.Forms.Control[] controls)
        {
            TextEdit textEdit;

            foreach (System.Windows.Forms.Control textBox in (TextEdit[])controls)
            {
                if (textBox is TextEdit)
                {
                    textEdit = (TextEdit)textBox;

                    if (this.FilterControl.ContainsKey(textEdit))
                    {
                        if (this.FilterControl[textEdit].AutoCompleteMode != AutoCompleteMode.None)
                        {
                            textEdit.TextChanged -= this.Control_TextChanged;
                            textEdit.MaskBox.AutoCompleteMode = AutoCompleteMode.None;
                            textEdit.MaskBox.AutoCompleteCustomSource = null;
                            textEdit.MaskBox.AutoCompleteSource = AutoCompleteSource.None;
                        }

                        this.FilterControl.Remove(textEdit);
                    }
                }
            }
        }

        void Control_TextChanged(object sender, EventArgs e)
        {
            GridView gridView;
            System.Data.DataView dataView;
            TextEdit textBox;
            FilterAttribute filterAttribute;
            StringBuilder stringBuilder;
            //string _Tmp;
            decimal decimalTmp;
            string text;

            textBox = (TextEdit)sender;
            filterAttribute = this.FilterControl[textBox];
            gridView = filterAttribute.GridView;

            if (gridView.DataSource == null)
                return;

            dataView = null;

            if (gridView.DataSource is System.Data.DataView)
                dataView = (System.Data.DataView)gridView.DataSource;

            if (gridView.DataSource is System.Data.DataSet)
                dataView = ((System.Data.DataSet)gridView.DataSource).Tables[gridView.GridControl.DataMember].DefaultView;

            if (gridView.DataSource is System.Data.DataTable)
                dataView = ((System.Data.DataTable)gridView.DataSource).DefaultView;

            if (dataView == null)
                return;

            text = textBox.Text;
            text = text.Replace("[", "[[").Replace("]", "]]").Replace("[[", "[[]").Replace("]]", "[]]").Replace("*", "[*]").Replace("%", "[%]").Replace("'", "''");

            stringBuilder = new StringBuilder();
            if (!text.Equals(""))
            {
                foreach (GridColumn _GridColumn in gridView.Columns)
                {
                    if (!_GridColumn.Visible)//보이는 컬럼만 검색
                        continue;

                    if (_GridColumn.Name == filterAttribute.ColumnName && !filterAttribute.IsSearchAll)
                    {
                        if (_GridColumn.AppearanceCell.TextOptions.HAlignment == HorzAlignment.Far)
                        {
                            if (text.ToTryDecimal(out decimalTmp))
                            {
                                if (filterAttribute.IsStartsWith)
                                    stringBuilder.AppendFormat("OR Convert([{0}], 'System.String') LIKE '{1}%' ", _GridColumn.FieldName, text);
                                //_Tmp += "OR Convert([" + _DataGridViewColumn.DataPropertyName + "], 'System.String') LIKE '" + _Text + "%' ";
                                else
                                    stringBuilder.AppendFormat("OR Convert([{0}], 'System.String') LIKE '%{1}%' ", _GridColumn.FieldName, text);
                                //_Tmp += "OR Convert([" + _DataGridViewColumn.DataPropertyName + "], 'System.String') LIKE '%" + _Text + "%' ";
                            }
                            else
                            {
                                if (filterAttribute.IsStartsWith)
                                    stringBuilder.AppendFormat("OR [{0}] LIKE '{1}%' ", _GridColumn.FieldName, text);
                                //_Tmp += "OR [" + _DataGridViewColumn.DataPropertyName + "] LIKE '" + _Text + "%' ";
                                else
                                    stringBuilder.AppendFormat("OR [{0}] LIKE '%{1}%' ", _GridColumn.FieldName, text);
                                //_Tmp += "OR [" + _DataGridViewColumn.DataPropertyName + "] LIKE '%" + _Text + "%' ";
                            }
                        }
                        else
                        {
                            if (filterAttribute.IsStartsWith)
                                stringBuilder.AppendFormat("OR [{0}] LIKE '{1}%' ", _GridColumn.FieldName, text);
                            //_Tmp += "OR [" + _DataGridViewColumn.DataPropertyName + "] LIKE '" + _Text + "%' ";
                            else
                                stringBuilder.AppendFormat("OR [{0}] LIKE '%{1}%' ", _GridColumn.FieldName, text);
                            //_Tmp += "OR [" + _DataGridViewColumn.DataPropertyName + "] LIKE '%" + _Text + "%' ";
                        }
                        break;
                    }

                    if (filterAttribute.IsSearchAll)
                    {
                        if (_GridColumn.AppearanceCell.TextOptions.HAlignment == HorzAlignment.Far)
                        {
                            if (text.ToTryDecimal(out decimalTmp))
                            {
                                if (filterAttribute.IsStartsWith)
                                    stringBuilder.AppendFormat("OR Convert([{0}], 'System.String') LIKE '{1}%' ", _GridColumn.FieldName, text);
                                //_Tmp += "OR Convert([" + _DataGridViewColumn.DataPropertyName + "], 'System.String') LIKE '" + _Text + "%' ";
                                else

                                    stringBuilder.AppendFormat("OR Convert([{0}], 'System.String') LIKE '%{1}%' ", _GridColumn.FieldName, text);
                                //_Tmp += "OR Convert([" + _DataGridViewColumn.DataPropertyName + "], 'System.String') LIKE '%" + _Text + "%' ";
                            }
                            else
                            {
                                if (filterAttribute.IsStartsWith)
                                    stringBuilder.AppendFormat("OR [{0}] LIKE '{1}%' ", _GridColumn.FieldName, text);
                                //_Tmp += "OR [" + _DataGridViewColumn.DataPropertyName + "] LIKE '" + _Text + "%' ";
                                else
                                    stringBuilder.AppendFormat("OR [{0}] LIKE '%{1}%' ", _GridColumn.FieldName, text);
                                //_Tmp += "OR [" + _DataGridViewColumn.DataPropertyName + "] LIKE '%" + _Text + "%' ";
                            }
                        }
                        else
                        {
                            if (filterAttribute.IsStartsWith)
                                stringBuilder.AppendFormat("OR [{0}] LIKE '{1}%' ", _GridColumn.FieldName, text);
                            //_Tmp += "OR [" + _DataGridViewColumn.DataPropertyName + "] LIKE '" + _Text + "%' ";
                            else
                                stringBuilder.AppendFormat("OR [{0}] LIKE '%{1}%' ", _GridColumn.FieldName, text);
                            //_Tmp += "OR [" + _DataGridViewColumn.DataPropertyName + "] LIKE '%" + _Text + "%' ";
                        }
                    }
                }
            }

            try
            {
                if (stringBuilder.ToString().StartsWith("OR "))
                    dataView.RowFilter = stringBuilder.ToString(3, stringBuilder.Length - 3);
                else
                    dataView.RowFilter = stringBuilder.ToString();
            }
            catch (Exception exception)
            {
                DiagnosticsTool.MyTrace(exception);
                dataView.RowFilter = "";
            }
        }

        private void GridView_DataSourceChanged(object sender, EventArgs e)
        {
            GridView gridView;
            System.Data.DataView dataView;

            gridView = (GridView)sender;

            if (gridView.DataSource == null)
                return;

            if (this.FilterControl == null)
                return;

            dataView = null;

            if (gridView.DataSource is System.Data.DataView)
                dataView = (System.Data.DataView)gridView.DataSource;

            if (gridView.DataSource is System.Data.DataSet)
                dataView = ((System.Data.DataSet)gridView.DataSource).Tables[gridView.GridControl.DataMember].DefaultView;

            if (gridView.DataSource is System.Data.DataTable)
                dataView = ((System.Data.DataTable)gridView.DataSource).DefaultView;

            if (dataView == null)
                return;

            foreach (FilterAttribute filterAttribute in this.FilterControl.Values)
            {
                //_FilterAttribute.AutoCompleteStringCollection.Clear();

                for (int i = 0; i < dataView.Table.Rows.Count; i++)
                {
                    if (dataView.Table.Columns.Contains(filterAttribute.ColumnName)
                        && !filterAttribute.AutoCompleteStringCollection.Contains(dataView.Table.Rows[i][filterAttribute.ColumnName].ToString()))
                        filterAttribute.AutoCompleteStringCollection.Add(dataView.Table.Rows[i][filterAttribute.ColumnName].ToString());
                }
            }

        }

        private void DataGridView_DataBindingComplete(object sender, EventArgs e)
        {
            GridView gridView;
            System.Data.DataView dataView;

            gridView = (GridView)sender;

            if (gridView.DataSource == null)
                return;

            dataView = null;

            if (gridView.DataSource is System.Data.DataView)
                dataView = (System.Data.DataView)gridView.DataSource;

            if (gridView.DataSource is System.Data.DataSet)
                dataView = ((System.Data.DataSet)gridView.DataSource).Tables[gridView.GridControl.DataMember].DefaultView;

            if (gridView.DataSource is System.Data.DataTable)
                dataView = ((System.Data.DataTable)gridView.DataSource).DefaultView;

            if (dataView == null)
                return;

            for (int i = 0; i <= gridView.RowCount - 1; i++)
            {
                if (gridView.IsNewItemRow(i)) continue;
            }

            CustomDrawRowIndicator(gridView, dataView);
            gridView.GridControl.MouseDoubleClick -= GridControl_MouseDoubleClick;
            gridView.GridControl.MouseDoubleClick += GridControl_MouseDoubleClick;
        }

        private void GridControl_MouseDoubleClick(object sender, MouseEventArgs e)
        {
            GridControl gridControl = (GridControl)sender;
            GridView gridView;

            gridView = (GridView)((GridControl)((IDataGridAgent)this).GridControl).MainView;

            GridHitInfo gridHitInfo = gridView.CalcHitInfo(e.Location);

            if (gridHitInfo.HitTest == GridHitTest.ColumnButton)
            {
                if (gridView.SelectedRowsCount == gridView.RowCount)
                    gridView.ClearSelection();
                else
                    gridView.SelectAll();
            }
        }

        private void CustomDrawRowIndicator(GridView gridView, System.Data.DataView dataView)
        {
            gridView.IndicatorWidth = 25 + (gridView.RowCount.ToString().Length * 14);

            gridView.CustomDrawRowIndicator += (s, e) => {
                if (e.Info.Kind == DevExpress.Utils.Drawing.IndicatorKind.Header)
                {
                    e.Info.DisplayText = gridView.RowCount.ToString() + "/" + dataView.Table.Rows.Count.ToString();
                    e.Appearance.TextOptions.HAlignment = HorzAlignment.Center;
                }
                else if (e.Info.Kind == DevExpress.Utils.Drawing.IndicatorKind.Row)
                {
                    if (e.RowHandle < 0)
                        e.Info.DisplayText = string.Empty;
                    else
                    {
                        e.Info.DisplayText = (e.RowHandle + 1).ToString();
                        e.Appearance.TextOptions.HAlignment = HorzAlignment.Far;
                    }
                }
            };
        }

        private void SetSkin(GridView gridView)
        {
            //string tmp;

            //try
            //{
            //    tmp = this.GetAttribute("SkinName");

            //    if (tmp != null || tmp != "")
            //    {
            //        gridView.GridControl.LookAndFeel.SkinName = tmp;
            //        gridView.GridControl.LookAndFeel.UseDefaultLookAndFeel = false;
            //    }
            //}
            //catch (Exception ex)
            //{
            //    DiagnosticsTool.MyTrace(ex);
            //}
        }

        protected struct FilterAttribute
        {
            public GridView GridView { get; set; }
            public string ColumnName { get; set; }
            public bool IsSearchAll { get; set; }
            public bool IsStartsWith { get; set; }
            public AutoCompleteMode AutoCompleteMode { get; set; }
            public AutoCompleteStringCollection AutoCompleteStringCollection { get; set; }
        }
    }
}