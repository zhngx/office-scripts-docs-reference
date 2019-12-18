export declare namespace ExcelScript {
    /*
     * Special Run Function
     */
    function run(
        callback: (workbook: Workbook) => Promise<void>
    ): Promise<void>;

    //
    // Class
    //

    /**
     * Represents the Excel application that manages the workbook.
     */
    interface Application {
        /**
         * Returns the Excel calculation engine version used for the last full recalculation. Read-only.
         */
        getCalculationEngineVersion(): number;

        /**
         * Returns the calculation mode used in the workbook, as defined by the constants in Excel.CalculationMode. Possible values are: `Automatic`, where Excel controls recalculation; `AutomaticExceptTables`, where Excel controls recalculation but ignores changes in tables; `Manual`, where calculation is done when the user requests it.
         */
        getCalculationMode(): CalculationMode;
        setCalculationMode(calculationMode: CalculationMode): void;

        /**
         * Returns the calculation state of the application. See Excel.CalculationState for details. Read-only.
         */
        getCalculationState(): CalculationState;

        /**
         * Returns the Iterative Calculation settings.
         * In Excel on Windows and Mac, the settings will apply to the Excel Application.
         * In Excel on the web and other platforms, the settings will apply to the active workbook.
         */
        getIterativeCalculation(): IterativeCalculation;

        /**
         * Recalculate all currently opened workbooks in Excel.
         */
        calculate(calculationType: CalculationType): void;
    }

    /**
     * Represents the Iterative Calculation settings.
     */
    interface IterativeCalculation {
        /**
         * True if Excel will use iteration to resolve circular references.
         */
        getEnabled(): boolean;
        setEnabled(enabled: boolean): void;

        /**
         * Returns or sets the maximum amount of change between each iteration as Excel resolves circular references.
         */
        getMaxChange(): number;
        setMaxChange(maxChange: number): void;

        /**
         * Returns or sets the maximum number of iterations that Excel can use to resolve a circular reference.
         */
        getMaxIteration(): number;
        setMaxIteration(maxIteration: number): void;
    }

    /**
     * Workbook is the top level object which contains related workbook objects such as worksheets, tables, ranges, etc.
     * To learn more about the workbook object model, read {@link https://docs.microsoft.com/office/dev/add-ins/excel/excel-add-ins-workbooks | Work with workbooks using the Excel JavaScript API}.
     */
    interface Workbook {
        /**
         * Represents the Excel application instance that contains this workbook. Read-only.
         */
        getApplication(): Application;

        /**
         * Specifies whether or not the workbook is in autosave mode. Read-Only.
         */
        getAutoSave(): boolean;

        /**
         * Returns a number about the version of Excel Calculation Engine. Read-Only.
         */
        getCalculationEngineVersion(): number;

        /**
         * True if all charts in the workbook are tracking the actual data points to which they are attached.
         * False if the charts track the index of the data points.
         */
        getChartDataPointTrack(): boolean;
        setChartDataPointTrack(chartDataPointTrack: boolean): void;

        /**
         * Specifies whether or not changes have been made since the workbook was last saved.
         * You can set this property to true if you want to close a modified workbook without either saving it or being prompted to save it.
         */
        getIsDirty(): boolean;
        setIsDirty(isDirty: boolean): void;

        /**
         * Gets the workbook name. Read-only.
         */
        getName(): string;

        /**
         * Specifies whether or not the workbook has ever been saved locally or online. Read-Only.
         */
        getPreviouslySaved(): boolean;

        /**
         * Gets the workbook properties. Read-only.
         */
        getProperties(): DocumentProperties;

        /**
         * Returns workbook protection object for a workbook. Read-only.
         */
        getProtection(): WorkbookProtection;

        /**
         * True if the workbook is open in Read-only mode. Read-only.
         */
        getReadOnly(): boolean;

        /**
         * True if calculations in this workbook will be done using only the precision of the numbers as they're displayed.
         * Data will permanently lose accuracy when switching this property from false to true.
         */
        getUsePrecisionAsDisplayed(): boolean;
        setUsePrecisionAsDisplayed(usePrecisionAsDisplayed: boolean): void;

        /**
         * Gets the currently active cell from the workbook.
         */
        getActiveCell(): Range;

        /**
         * Gets the currently active chart in the workbook. If there is no active chart, a null object is returned.
         */
        getActiveChart(): Chart;

        /**
         * Gets the currently active slicer in the workbook. If there is no active slicer, a null object is returned.
         */
        getActiveSlicer(): Slicer;

        /**
         * Gets the currently selected single range from the workbook. If there are multiple ranges selected, this method will throw an error.
         */
        getSelectedRange(): Range;

        /**
         * Gets the currently selected one or more ranges from the workbook. Unlike getSelectedRange(), this method returns a RangeAreas object that represents all the selected ranges.
         */
        getSelectedRanges(): RangeAreas;

        getBindings(): Binding[];
        addBinding(
            range: Range | string,
            bindingType: BindingType,
            id: string
        ): Binding;
        addBindingFromNamedItem(
            name: string,
            bindingType: BindingType,
            id: string
        ): Binding;
        addBindingFromSelection(bindingType: BindingType, id: string): Binding;
        getBinding(id: string): Binding | undefined;

        getComments(): Comment[];
        addComment(
            cellAddress: Range | string,
            content: CommentRichContent | string,
            contentType?: ContentType
        ): Comment;
        getComment(commentId: string): Comment;
        getCommentByCell(cellAddress: Range | string): Comment;
        getCommentByReplyId(replyId: string): Comment;

        getCustomXmlParts(): CustomXmlPart[];
        addCustomXmlPart(xml: string): CustomXmlPart;
        getCustomXmlPart(id: string): CustomXmlPart | undefined;

        getNames(): NamedItem[];
        addNamedItem(
            name: string,
            reference: Range | string,
            comment?: string
        ): NamedItem;
        addNamedItemFormulaLocal(
            name: string,
            formula: string,
            comment?: string
        ): NamedItem;
        getNamedItem(name: string): NamedItem | undefined;

        getPivotTableStyles(): PivotTableStyle[];
        addPivotTableStyle(
            name: string,
            makeUniqueName?: boolean
        ): PivotTableStyle;
        getDefaultPivotTableStyle(): PivotTableStyle;
        getPivotTableStyle(name: string): PivotTableStyle | undefined;
        setDefaultPivotTableStyle(
            newDefaultStyle: PivotTableStyle | string
        ): void;

        getPivotTables(): PivotTable[];
        addPivotTable(
            name: string,
            source: Range | string | Table,
            destination: Range | string
        ): PivotTable;
        getPivotTable(name: string): PivotTable | undefined;
        refreshAllPivotTables(): void;

        getSettings(): Setting[];
        addSetting(
            key: string,
            value: string | number | boolean | Date | Array<any> | any
        ): Setting;
        getSetting(key: string): Setting | undefined;

        getSlicerStyles(): SlicerStyle[];
        addSlicerStyle(name: string, makeUniqueName?: boolean): SlicerStyle;
        getDefaultSlicerStyle(): SlicerStyle;
        getSlicerStyle(name: string): SlicerStyle | undefined;
        setDefaultSlicerStyle(newDefaultStyle: SlicerStyle | string): void;

        getSlicers(): Slicer[];
        addSlicer(
            slicerSource: string | PivotTable | Table,
            sourceField: string | PivotField | number | TableColumn,
            slicerDestination?: string | Worksheet
        ): Slicer;
        getSlicer(key: string): Slicer | undefined;

        getStyles(): Style[];
        addvoid(name: string): void;
        getStyle(name: string): Style;

        getTableStyles(): TableStyle[];
        addTableStyle(name: string, makeUniqueName?: boolean): TableStyle;
        getDefaultTableStyle(): TableStyle;
        getTableStyle(name: string): TableStyle | undefined;
        setDefaultTableStyle(newDefaultStyle: TableStyle | string): void;

        getTables(): Table[];
        addTable(address: Range | string, hasHeaders: boolean): Table;
        getTable(key: string): Table | undefined;

        getTimelineStyles(): TimelineStyle[];
        addTimelineStyle(name: string, makeUniqueName?: boolean): TimelineStyle;
        getDefaultTimelineStyle(): TimelineStyle;
        getTimelineStyle(name: string): TimelineStyle | undefined;
        setDefaultTimelineStyle(newDefaultStyle: TimelineStyle | string): void;

        getWorksheets(): Worksheet[];
        addWorksheet(name?: string): Worksheet;
        getActiveWorksheet(): Worksheet;
        getFirstWorksheet(visibleOnly?: boolean): Worksheet;
        getWorksheet(key: string): Worksheet | undefined;
        getLastWorksheet(visibleOnly?: boolean): Worksheet;
    }

    /**
     * Represents the protection of a workbook object.
     */
    interface WorkbookProtection {
        /**
         * Indicates if the workbook is protected. Read-Only.
         */
        getProtected(): boolean;

        /**
         * Protects a workbook. Fails if the workbook has been protected.
         */
        protect(password?: string): void;

        /**
         * Unprotects a workbook.
         */
        unprotect(password?: string): void;
    }

    /**
     * An Excel worksheet is a grid of cells. It can contain data, tables, charts, etc.
     * To learn more about the worksheet object model, read {@link https://docs.microsoft.com/office/dev/add-ins/excel/excel-add-ins-worksheets | Work with worksheets using the Excel JavaScript API}.
     */
    interface Worksheet {
        /**
         * Represents the AutoFilter object of the worksheet. Read-Only.
         */
        getAutoFilter(): AutoFilter;

        /**
         * Gets or sets the enableCalculation property of the worksheet.
         * True if Excel recalculates the worksheet when necessary. False if Excel doesn't recalculate the sheet.
         */
        getEnableCalculation(): boolean;
        setEnableCalculation(enableCalculation: boolean): void;

        /**
         * Gets an object that can be used to manipulate frozen panes on the worksheet. Read-only.
         */
        getFreezePanes(): WorksheetFreezePanes;

        /**
         * Returns a value that uniquely identifies the worksheet in a given workbook. The value of the identifier remains the same even when the worksheet is renamed or moved. Read-only.
         */
        getId(): string;

        /**
         * The display name of the worksheet.
         */
        getName(): string;
        setName(name: string): void;

        /**
         * Gets the PageLayout object of the worksheet.
         */
        getPageLayout(): PageLayout;

        /**
         * The zero-based position of the worksheet within the workbook.
         */
        getPosition(): number;
        setPosition(position: number): void;

        /**
         * Returns sheet protection object for a worksheet. Read-only.
         */
        getProtection(): WorksheetProtection;

        /**
         * Gets or sets the worksheet's gridlines flag.
         * This flag determines whether gridlines are visible to the user.
         */
        getShowGridlines(): boolean;
        setShowGridlines(showGridlines: boolean): void;

        /**
         * Gets or sets the worksheet's headings flag.
         * This flag determines whether headings are visible to the user.
         */
        getShowHeadings(): boolean;
        setShowHeadings(showHeadings: boolean): void;

        /**
         * Returns the standard (default) height of all the rows in the worksheet, in points. Read-only.
         */
        getStandardHeight(): number;

        /**
         * Returns or sets the standard (default) width of all the columns in the worksheet.
         * One unit of column width is equal to the width of one character in the Normal style. For proportional fonts, the width of the character 0 (zero) is used.
         */
        getStandardWidth(): number;
        setStandardWidth(standardWidth: number): void;

        /**
         * Gets or sets the worksheet tab color.
         * When retrieving the tab color, if the worksheet is invisible, the value will be null. If the worksheet is visible but the tab color is set to auto, an empty string will be returned. Otherwise, the property will be set to a color, in the form "#123456"
         * When setting the color, use an empty-string to set an "auto" color, or a real color otherwise.
         */
        getTabColor(): string;
        setTabColor(tabColor: string): void;

        /**
         * The Visibility of the worksheet.
         */
        getVisibility(): SheetVisibility;
        setVisibility(visibility: SheetVisibility): void;

        /**
         * Activate the worksheet in the Excel UI.
         */
        activate(): void;

        /**
         * Calculates all cells on a worksheet.
         */
        calculate(markAllDirty: boolean): void;

        /**
         * Copies a worksheet and places it at the specified position.
         */
        copy(
            positionType?: WorksheetPositionType,
            relativeTo?: Worksheet
        ): Worksheet;

        /**
         * Deletes the worksheet from the workbook. Note that if the worksheet's visibility is set to "VeryHidden", the delete operation will fail with an `InvalidOperation` exception. You should first change its visibility to hidden or visible before deleting it.
         */
        delete(): void;

        /**
         * Finds all occurrences of the given string based on the criteria specified and returns them as a RangeAreas object, comprising one or more rectangular ranges.
         */
        findAll(text: string, criteria: WorksheetSearchCriteria): RangeAreas;

        /**
         * Gets the range object containing the single cell based on row and column numbers. The cell can be outside the bounds of its parent range, so long as it stays within the worksheet grid.
         */
        getCell(row: number, column: number): Range;

        /**
         * Gets the worksheet that follows this one. If there are no worksheets following this one, this method will return a null object.
         */
        getNext(visibleOnly?: boolean): Worksheet;

        /**
         * Gets the worksheet that precedes this one. If there are no previous worksheets, this method will return a null objet.
         */
        getPrevious(visibleOnly?: boolean): Worksheet;

        /**
         * Gets the range object, representing a single rectangular block of cells, specified by the address or name.
         */
        getRange(address?: string): Range;

        /**
         * Gets the range object beginning at a particular row index and column index, and spanning a certain number of rows and columns.
         */
        getRangeByIndexes(
            startRow: number,
            startColumn: number,
            rowCount: number,
            columnCount: number
        ): Range;

        /**
         * Gets the RangeAreas object, representing one or more blocks of rectangular ranges, specified by the address or name.
         */
        getRanges(address?: string): RangeAreas;

        /**
         * The used range is the smallest range that encompasses any cells that have a value or formatting assigned to them. If the entire worksheet is blank, this function will return a null object.
         */
        getUsedRange(valuesOnly?: boolean): Range;

        /**
         * Finds and replaces the given string based on the criteria specified within the current worksheet.
         */
        replaceAll(
            text: string,
            replacement: string,
            criteria: ReplaceCriteria
        ): number;

        /**
         * Shows row or column groups by their outline levels.
         * Outlines group and summarize a list of data in the worksheet.
         * The `rowLevels` and `columnLevels` parameters specify how many levels of the outline will be displayed.
         *  The acceptable argument range is between 0 and 8.
         *  A value of 0 does not change the current display. A value greater than the current number of levels displays all the levels.
         */
        showOutlineLevels(rowLevels: number, columnLevels: number): void;

        getCharts(): Chart[];
        addChart(
            type: ChartType,
            sourceData: Range,
            seriesBy?: ChartSeriesBy
        ): Chart;
        getChart(name: string): Chart | undefined;

        getComments(): Comment[];
        addComment(
            cellAddress: Range | string,
            content: CommentRichContent | string,
            contentType?: ContentType
        ): Comment;
        getComment(commentId: string): Comment;
        getCommentByCell(cellAddress: Range | string): Comment;
        getCommentByReplyId(replyId: string): Comment;

        getNames(): NamedItem[];
        addNamedItem(
            name: string,
            reference: Range | string,
            comment?: string
        ): NamedItem;
        addNamedItemFormulaLocal(
            name: string,
            formula: string,
            comment?: string
        ): NamedItem;
        getNamedItem(name: string): NamedItem | undefined;

        getPivotTables(): PivotTable[];
        addPivotTable(
            name: string,
            source: Range | string | Table,
            destination: Range | string
        ): PivotTable;
        getPivotTable(name: string): PivotTable | undefined;
        refreshAllPivotTables(): void;

        getShapes(): Shape[];
        addGeometricShape(geometricShapeType: GeometricShapeType): Shape;
        addGroup(values: Array<string | Shape>): Shape;
        addImage(base64ImageString: string): Shape;
        addLine(
            startLeft: number,
            startTop: number,
            endLeft: number,
            endTop: number,
            connectorType?: ConnectorType
        ): Shape;
        addTextBox(text?: string): Shape;
        getShape(key: string): Shape;

        getSlicers(): Slicer[];
        addSlicer(
            slicerSource: string | PivotTable | Table,
            sourceField: string | PivotField | number | TableColumn,
            slicerDestination?: string | Worksheet
        ): Slicer;
        getSlicer(key: string): Slicer | undefined;

        getTables(): Table[];
        addTable(address: Range | string, hasHeaders: boolean): Table;
        getTable(key: string): Table | undefined;
    }

    /**
     * Represents the protection of a sheet object.
     */
    interface WorksheetProtection {
        /**
         * Sheet protection options. Read-only.
         */
        getOptions(): WorksheetProtectionOptions;

        /**
         * Indicates if the worksheet is protected. Read-only.
         */
        getProtected(): boolean;

        /**
         * Protects a worksheet. Fails if the worksheet has already been protected.
         */
        protect(options?: WorksheetProtectionOptions, password?: string): void;

        /**
         * Unprotects a worksheet.
         */
        unprotect(password?: string): void;
    }

    /**
     * no comment
     */
    interface WorksheetFreezePanes {
        /**
         * Sets the frozen cells in the active worksheet view.
         * The range provided corresponds to cells that will be frozen in the top- and left-most pane.
         */
        freezeAt(frozenRange: Range | string): void;

        /**
         * Freeze the first column(s) of the worksheet in place.
         */
        freezeColumns(count?: number): void;

        /**
         * Freeze the top row(s) of the worksheet in place.
         */
        freezeRows(count?: number): void;

        /**
         * Gets a range that describes the frozen cells in the active worksheet view.
         * The frozen range is corresponds to cells that are frozen in the top- and left-most pane.
         * If there is no frozen pane, returns a null object.
         */
        getLocation(): Range;

        /**
         * Removes all frozen panes in the worksheet.
         */
        unfreeze(): void;
    }

    /**
     * Range represents a set of one or more contiguous cells such as a cell, a row, a column, block of cells, etc.
     *  To learn more about how ranges are used throughout the API, read {@link https://docs.microsoft.com/office/dev/add-ins/excel/excel-add-ins-ranges | Work with ranges using the Excel JavaScript API}
     *  and {@link https://docs.microsoft.com/office/dev/add-ins/excel/excel-add-ins-ranges-advanced | Work with ranges using the Excel JavaScript API (advanced)}.
     */
    interface Range {
        /**
         * Represents the range reference in A1-style. Address value will contain the Sheet reference (e.g. "Sheet1!A1:B4"). Read-only.
         */
        getAddress(): string;

        /**
         * Represents range reference for the specified range in the language of the user. Read-only.
         */
        getAddressLocal(): string;

        /**
         * Number of cells in the range. This API will return -1 if the cell count exceeds 2^31-1 (2,147,483,647). Read-only.
         */
        getCellCount(): number;

        /**
         * Represents the total number of columns in the range. Read-only.
         */
        getColumnCount(): number;

        /**
         * Represents if all columns of the current range are hidden.
         */
        getColumnHidden(): boolean;
        setColumnHidden(columnHidden: boolean): void;

        /**
         * Represents the column number of the first cell in the range. Zero-indexed. Read-only.
         */
        getColumnIndex(): number;

        /**
         * Returns a data validation object.
         */
        getDataValidation(): DataValidation;

        /**
         * Returns a format object, encapsulating the range's font, fill, borders, alignment, and other properties. Read-only.
         */
        getFormat(): RangeFormat;

        /**
         * Returns the distance in points, for 100% zoom, from top edge of the range to bottom edge of the range. Read-only.
         */
        getHeight(): number;

        /**
         * Represents if all cells of the current range are hidden. Read-only.
         */
        getHidden(): boolean;

        /**
         * Represents the hyperlink for the current range.
         */
        getHyperlink(): RangeHyperlink;
        setHyperlink(hyperlink: RangeHyperlink): void;

        /**
         * Represents if the current range is an entire column. Read-only.
         */
        getIsEntireColumn(): boolean;

        /**
         * Represents if the current range is an entire row. Read-only.
         */
        getIsEntireRow(): boolean;

        /**
         * Returns the distance in points, for 100% zoom, from left edge of the worksheet to left edge of the range. Read-only.
         */
        getLeft(): number;

        /**
         * Returns the total number of rows in the range. Read-only.
         */
        getRowCount(): number;

        /**
         * Represents if all rows of the current range are hidden.
         */
        getRowHidden(): boolean;
        setRowHidden(rowHidden: boolean): void;

        /**
         * Returns the row number of the first cell in the range. Zero-indexed. Read-only.
         */
        getRowIndex(): number;

        /**
         * Represents the range sort of the current range. Read-only.
         */
        getSort(): RangeSort;

        /**
         * Represents the style of the current range.
         * If the styles of the cells are inconsistent, null will be returned.
         * For custom styles, the style name will be returned. For built-in styles, a string representing a value in the BuiltInStyle enum will be returned.
         */
        getStyle(): string;
        setStyle(style: string): void;

        /**
         * Text values of the specified range. The Text value will not depend on the cell width. The # sign substitution that happens in Excel UI will not affect the text value returned by the API. Read-only.
         */
        getText(): string[][];

        /**
         * Returns the distance in points, for 100% zoom, from top edge of the worksheet to top edge of the range. Read-only.
         */
        getTop(): number;

        /**
         * Represents the raw values of the specified range. The data returned could be of type string, number, or a boolean. Cells that contain an error will return the error string.
         * When setting values to a range, the value argument can be either a single value (string, number or boolean) or a two-dimensional array. If the argument is a single value, it will be applied to all cells in the range.
         */
        getValues(): any[][];
        setValues(values: any[][]): void;

        /**
         * Returns the distance in points, for 100% zoom, from left edge of the range to right edge of the range. Read-only.
         */
        getWidth(): number;

        /**
         * The worksheet containing the current range. Read-only.
         */
        getWorksheet(): Worksheet;

        /**
         * Fills range from the current range to the destination range using the specified AutoFill logic.
         *  The destination range can be null, or can extend the source either horizontally or vertically.
         *  Discontiguous ranges are not supported.
         *
         *  For more information, read {@link https://support.office.com/article/video-use-autofill-and-flash-fill-2e79a709-c814-4b27-8bc2-c4dc84d49464 | Use AutoFill and Flash Fill}.
         */
        autoFill(
            destinationRange?: Range | string,
            autoFillType?: AutoFillType
        ): void;

        /**
         * Calculates a range of cells on a worksheet.
         */
        calculate(): void;

        /**
         * Clear range values, format, fill, border, etc.
         */
        clear(applyTo?: ClearApplyTo): void;

        /**
         * Converts the range cells with datatypes into text.
         */
        convertDataTypeToText(): void;

        /**
         * Converts the range cells into linked datatype in the worksheet.
         */
        convertToLinkedDataType(
            serviceID: number,
            languageCulture: string
        ): void;

        /**
         * Copies cell data or formatting from the source range or RangeAreas to the current range.
         * The destination range can be a different size than the source range or RangeAreas. The destination will be expanded automatically if it is smaller than the source.
         */
        copyFrom(
            sourceRange: Range | RangeAreas | string,
            copyType?: RangeCopyType,
            skipBlanks?: boolean,
            transpose?: boolean
        ): void;

        /**
         * Deletes the cells associated with the range.
         */
        delete(shift: DeleteShiftDirection): void;

        /**
         * Finds the given string based on the criteria specified.
         * If the current range is larger than a single cell, then the search will be limited to that range, else the search will cover the entire sheet starting after that cell.
         * If there are no matches, this function will return a null object.
         */
        find(text: string, criteria: SearchCriteria): Range;

        /**
         * Does FlashFill to current range.Flash Fill will automatically fills data when it senses a pattern, so the range must be single column range and have data around in order to find pattern.
         */
        flashFill(): void;

        /**
         * Gets a Range object with the same top-left cell as the current Range object, but with the specified numbers of rows and columns.
         */
        getAbsoluteResizedRange(numRows: number, numColumns: number): Range;

        /**
         * Gets the smallest range object that encompasses the given ranges. For example, the GetBoundingRect of "B2:C5" and "D10:E15" is "B2:E15".
         */
        getBoundingRect(anotherRange: Range | string): Range;

        /**
         * Gets the range object containing the single cell based on row and column numbers. The cell can be outside the bounds of its parent range, so long as it stays within the worksheet grid. The returned cell is located relative to the top left cell of the range.
         */
        getCell(row: number, column: number): Range;

        /**
         * Gets a column contained in the range.
         */
        getColumn(column: number): Range;

        /**
         * Gets a certain number of columns to the right of the current Range object.
         */
        getColumnsAfter(count?: number): Range;

        /**
         * Gets a certain number of columns to the left of the current Range object.
         */
        getColumnsBefore(count?: number): Range;

        /**
         * Gets an object that represents the entire column of the range (for example, if the current range represents cells "B4:E11", its `getEntireColumn` is a range that represents columns "B:E").
         */
        getEntireColumn(): Range;

        /**
         * Gets an object that represents the entire row of the range (for example, if the current range represents cells "B4:E11", its `GetEntireRow` is a range that represents rows "4:11").
         */
        getEntireRow(): Range;

        /**
         * Renders the range as a base64-encoded png image.
         */
        getImage(): string;

        /**
         * Gets the range object that represents the rectangular intersection of the given ranges. If no intersection is found, will return a null object.
         */
        getIntersection(anotherRange: Range | string): Range;

        /**
         * Gets the last cell within the range. For example, the last cell of "B2:D5" is "D5".
         */
        getLastCell(): Range;

        /**
         * Gets the last column within the range. For example, the last column of "B2:D5" is "D2:D5".
         */
        getLastColumn(): Range;

        /**
         * Gets the last row within the range. For example, the last row of "B2:D5" is "B5:D5".
         */
        getLastRow(): Range;

        /**
         * Gets an object which represents a range that's offset from the specified range. The dimension of the returned range will match this range. If the resulting range is forced outside the bounds of the worksheet grid, an error will be thrown.
         */
        getOffsetRange(rowOffset: number, columnOffset: number): Range;

        /**
         * Gets a Range object similar to the current Range object, but with its bottom-right corner expanded (or contracted) by some number of rows and columns.
         */
        getResizedRange(deltaRows: number, deltaColumns: number): Range;

        /**
         * Gets a row contained in the range.
         */
        getRow(row: number): Range;

        /**
         * Gets a certain number of rows above the current Range object.
         */
        getRowsAbove(count?: number): Range;

        /**
         * Gets a certain number of rows below the current Range object.
         */
        getRowsBelow(count?: number): Range;

        /**
         * Gets the RangeAreas object, comprising one or more ranges, that represents all the cells that match the specified type and value.
         * If no special cells are found, a null object will be returned.
         */
        getSpecialCells(
            cellType: SpecialCellType,
            cellValueType?: SpecialCellValueType
        ): RangeAreas;

        /**
         * Returns a Range object that represents the surrounding region for the top-left cell in this range. A surrounding region is a range bounded by any combination of blank rows and blank columns relative to this range.
         */
        getSurroundingRegion(): Range;

        /**
         * Returns the used range of the given range object. If there are no used cells within the range, this function will return a null object.
         */
        getUsedRange(valuesOnly?: boolean): Range;

        /**
         * Represents the visible rows of the current range.
         */
        getVisibleView(): RangeView;

        /**
         * Groups columns and rows for an outline.
         */
        group(groupOption: GroupOption): void;

        /**
         * Hide details of the row or column group.
         */
        hideGroupDetails(groupOption: GroupOption): void;

        /**
         * Inserts a cell or a range of cells into the worksheet in place of this range, and shifts the other cells to make space. Returns a new Range object at the now blank space.
         */
        insert(shift: InsertShiftDirection): Range;

        /**
         * Merge the range cells into one region in the worksheet.
         */
        merge(across?: boolean): void;

        /**
         * Removes duplicate values from the range specified by the columns.
         */
        removeDuplicates(
            columns: number[],
            includesHeader: boolean
        ): RemoveDuplicatesResult;

        /**
         * Finds and replaces the given string based on the criteria specified within the current range.
         */
        replaceAll(
            text: string,
            replacement: string,
            criteria: ReplaceCriteria
        ): number;

        /**
         * Selects the specified range in the Excel UI.
         */
        select(): void;

        /**
         * Set a range to be recalculated when the next recalculation occurs.
         */
        setDirty(): void;

        /**
         * Displays the card for an active cell if it has rich value content.
         */
        showCard(): void;

        /**
         * Show details of the row or column group.
         */
        showGroupDetails(groupOption: GroupOption): void;

        /**
         * Ungroups columns and rows for an outline.
         */
        ungroup(groupOption: GroupOption): void;

        /**
         * Unmerge the range cells into separate cells.
         */
        unmerge(): void;

        getConditionalFormats(): ConditionalFormat[];
        addConditionalFormat(type: ConditionalFormatType): ConditionalFormat;
        getConditionalFormat(id: string): ConditionalFormat;
    }

    /**
     * RangeAreas represents a collection of one or more rectangular ranges in the same worksheet.
     * To learn how to use discontinguous ranges, read {@link https://docs.microsoft.com/office/dev/add-ins/excel/excel-add-ins-multiple-ranges | Work with multiple ranges simultaneously in Excel add-ins}.
     */
    interface RangeAreas {
        /**
         * Returns the RageAreas reference in A1-style. Address value will contain the worksheet name for each rectangular block of cells (e.g. "Sheet1!A1:B4, Sheet1!D1:D4"). Read-only.
         */
        getAddress(): string;

        /**
         * Returns the RageAreas reference in the user locale. Read-only.
         */
        getAddressLocal(): string;

        /**
         * Returns the number of rectangular ranges that comprise this RangeAreas object.
         */
        getAreaCount(): number;

        /**
         * Returns the number of cells in the RangeAreas object, summing up the cell counts of all of the individual rectangular ranges. Returns -1 if the cell count exceeds 2^31-1 (2,147,483,647). Read-only.
         */
        getCellCount(): number;

        /**
         * Returns a dataValidation object for all ranges in the RangeAreas.
         */
        getDataValidation(): DataValidation;

        /**
         * Returns a rangeFormat object, encapsulating the the font, fill, borders, alignment, and other properties for all ranges in the RangeAreas object. Read-only.
         */
        getFormat(): RangeFormat;

        /**
         * Indicates whether all the ranges on this RangeAreas object represent entire columns (e.g., "A:C, Q:Z"). Read-only.
         */
        getIsEntireColumn(): boolean;

        /**
         * Indicates whether all the ranges on this RangeAreas object represent entire rows (e.g., "1:3, 5:7"). Read-only.
         */
        getIsEntireRow(): boolean;

        /**
         * Represents the style for all ranges in this RangeAreas object.
         * If the styles of the cells are inconsistent, null will be returned.
         * For custom styles, the style name will be returned. For built-in styles, a string representing a value in the BuiltInStyle enum will be returned.
         */
        getStyle(): string;
        setStyle(style: string): void;

        /**
         * Returns the worksheet for the current RangeAreas. Read-only.
         */
        getWorksheet(): Worksheet;

        /**
         * Calculates all cells in the RangeAreas.
         */
        calculate(): void;

        /**
         * Clears values, format, fill, border, etc on each of the areas that comprise this RangeAreas object.
         */
        clear(applyTo?: ClearApplyTo): void;

        /**
         * Converts all cells in the RangeAreas with datatypes into text.
         */
        convertDataTypeToText(): void;

        /**
         * Converts all cells in the RangeAreas into linked datatype.
         */
        convertToLinkedDataType(
            serviceID: number,
            languageCulture: string
        ): void;

        /**
         * Copies cell data or formatting from the source range or RangeAreas to the current RangeAreas.
         * The destination rangeAreas can be a different size than the source range or RangeAreas. The destination will be expanded automatically if it is smaller than the source.
         */
        copyFrom(
            sourceRange: Range | RangeAreas | string,
            copyType?: RangeCopyType,
            skipBlanks?: boolean,
            transpose?: boolean
        ): void;

        /**
         * Returns a RangeAreas object that represents the entire columns of the RangeAreas (for example, if the current RangeAreas represents cells "B4:E11, H2", it returns a RangeAreas that represents columns "B:E, H:H").
         */
        getEntireColumn(): RangeAreas;

        /**
         * Returns a RangeAreas object that represents the entire rows of the RangeAreas (for example, if the current RangeAreas represents cells "B4:E11", it returns a RangeAreas that represents rows "4:11").
         */
        getEntireRow(): RangeAreas;

        /**
         * Returns the RangeAreas object that represents the intersection of the given ranges or RangeAreas. If no intersection is found, a null object is returned.
         */
        getIntersection(anotherRange: Range | RangeAreas | string): RangeAreas;

        /**
         * Returns an RangeAreas object that is shifted by the specific row and column offset. The dimension of the returned RangeAreas will match the original object. If the resulting RangeAreas is forced outside the bounds of the worksheet grid, an error will be thrown.
         */
        getOffsetRangeAreas(
            rowOffset: number,
            columnOffset: number
        ): RangeAreas;

        /**
         * Returns a RangeAreas object that represents all the cells that match the specified type and value. Returns a null object if no special cells are found that match the criteria.
         */
        getSpecialCells(
            cellType: SpecialCellType,
            cellValueType?: SpecialCellValueType
        ): RangeAreas;

        /**
         * Returns the used RangeAreas that comprises all the used areas of individual rectangular ranges in the RangeAreas object.
         * If there are no used cells within the RangeAreas, a null object will be returned.
         */
        getUsedRangeAreas(valuesOnly?: boolean): RangeAreas;

        /**
         * Sets the RangeAreas to be recalculated when the next recalculation occurs.
         */
        setDirty(): void;

        getAreas(): Range[];

        getConditionalFormats(): ConditionalFormat[];
        addConditionalFormat(type: ConditionalFormatType): ConditionalFormat;
        getConditionalFormat(id: string): ConditionalFormat;
    }

    /**
     * RangeView represents a set of visible cells of the parent range.
     */
    interface RangeView {
        /**
         * Returns the number of visible columns. Read-only.
         */
        getColumnCount(): number;

        /**
         * Returns a value that represents the index of the RangeView. Read-only.
         */
        getIndex(): number;

        /**
         * Returns the number of visible rows. Read-only.
         */
        getRowCount(): number;

        /**
         * Text values of the specified range. The Text value will not depend on the cell width. The # sign substitution that happens in Excel UI will not affect the text value returned by the API. Read-only.
         */
        getText(): string[][];

        /**
         * Gets the parent range associated with the current RangeView.
         */
        getRange(): Range;

        getRows(): RangeView[];
    }

    /**
     * Setting represents a key-value pair of a setting persisted to the document (per file per add-in). These custom key-value pair can be used to store state or lifecycle information needed by the content or task-pane add-in. Note that settings are persisted in the document and hence it is not a place to store any sensitive or protected information such as user information and password.
     */
    interface Setting {
        /**
         * Returns the key that represents the id of the Setting. Read-only.
         */
        getKey(): string;

        /**
         * Represents the value stored for this setting.
         */
        getValue(): any;
        setValue(value: any): void;

        /**
         * Deletes the setting.
         */
        delete(): void;
    }

    /**
     * Represents a defined name for a range of cells or value. Names can be primitive named objects (as seen in the type below), range object, or a reference to a range. This object can be used to obtain range object associated with names.
     */
    interface NamedItem {
        /**
         * Returns an object containing values and types of the named item. Read-only.
         */
        getArrayValues(): NamedItemArrayValues;

        /**
         * Represents the comment associated with this name.
         */
        getComment(): string;
        setComment(comment: string): void;

        /**
         * The name of the object. Read-only.
         */
        getName(): string;

        /**
         * Indicates whether the name is scoped to the workbook or to a specific worksheet. Possible values are: Worksheet, Workbook. Read-only.
         */
        getScope(): NamedItemScope;

        /**
         * Indicates the type of the value returned by the name's formula. See Excel.NamedItemType for details. Read-only.
         */
        getType(): NamedItemType;

        /**
         * Specifies whether the object is visible or not.
         */
        getVisible(): boolean;
        setVisible(visible: boolean): void;

        /**
         * Deletes the given name.
         */
        delete(): void;

        /**
         * Returns the range object that is associated with the name. Returns a null object if the named item's type is not a range.
         */
        getRange(): Range;
    }

    /**
     * Represents an object containing values and types of a named item.
     */
    interface NamedItemArrayValues {
        /**
         * Represents the values of each item in the named item array.
         */
        getValues(): any[][];
    }

    /**
     * Represents an Office.js binding that is defined in the workbook.
     */
    interface Binding {
        /**
         * Represents binding identifier. Read-only.
         */
        getId(): string;

        /**
         * Returns the type of the binding. See Excel.BindingType for details. Read-only.
         */
        getType(): BindingType;

        /**
         * Deletes the binding.
         */
        delete(): void;

        /**
         * Returns the range represented by the binding. Will throw an error if binding is not of the correct type.
         */
        getRange(): Range;

        /**
         * Returns the table represented by the binding. Will throw an error if binding is not of the correct type.
         */
        getTable(): Table;

        /**
         * Returns the text represented by the binding. Will throw an error if binding is not of the correct type.
         */
        getText(): string;
    }

    /**
     * Represents an Excel table.
     * To learn more about the table object model, read {@link https://docs.microsoft.com/office/dev/add-ins/excel/excel-add-ins-tables | Work with tables using the Excel JavaScript API}.
     */
    interface Table {
        /**
         * Represents the AutoFilter object of the table. Read-Only.
         */
        getAutoFilter(): AutoFilter;

        /**
         * Indicates whether the first column contains special formatting.
         */
        getHighlightFirstColumn(): boolean;
        setHighlightFirstColumn(highlightFirstColumn: boolean): void;

        /**
         * Indicates whether the last column contains special formatting.
         */
        getHighlightLastColumn(): boolean;
        setHighlightLastColumn(highlightLastColumn: boolean): void;

        /**
         * Returns a value that uniquely identifies the table in a given workbook. The value of the identifier remains the same even when the table is renamed. Read-only.
         */
        getId(): string;

        /**
         * Returns a numeric id.
         */
        getLegacyId(): string;

        /**
         * Name of the table.
         *
         *   The set name of the table must follow the guidelines specified in the {@link https://support.office.com/article/Rename-an-Excel-table-FBF49A4F-82A3-43EB-8BA2-44D21233B114 | Rename an Excel table} article.
         */
        getName(): string;
        setName(name: string): void;

        /**
         * Indicates whether the columns show banded formatting in which odd columns are highlighted differently from even ones to make reading the table easier.
         */
        getShowBandedColumns(): boolean;
        setShowBandedColumns(showBandedColumns: boolean): void;

        /**
         * Indicates whether the rows show banded formatting in which odd rows are highlighted differently from even ones to make reading the table easier.
         */
        getShowBandedRows(): boolean;
        setShowBandedRows(showBandedRows: boolean): void;

        /**
         * Indicates whether the filter buttons are visible at the top of each column header. Setting this is only allowed if the table contains a header row.
         */
        getShowFilterButton(): boolean;
        setShowFilterButton(showFilterButton: boolean): void;

        /**
         * Indicates whether the header row is visible or not. This value can be set to show or remove the header row.
         */
        getShowHeaders(): boolean;
        setShowHeaders(showHeaders: boolean): void;

        /**
         * Indicates whether the total row is visible or not. This value can be set to show or remove the total row.
         */
        getShowTotals(): boolean;
        setShowTotals(showTotals: boolean): void;

        /**
         * Represents the sorting for the table. Read-only.
         */
        getSort(): TableSort;

        /**
         * Constant value that represents the Table style. Possible values are: "TableStyleLight1" through "TableStyleLight21", "TableStyleMedium1" through "TableStyleMedium28", "TableStyleDark1" through "TableStyleDark11". A custom user-defined style present in the workbook can also be specified.
         */
        getStyle(): string;
        setStyle(style: string): void;

        /**
         * The worksheet containing the current table. Read-only.
         */
        getWorksheet(): Worksheet;

        /**
         * Clears all the filters currently applied on the table.
         */
        clearFilters(): void;

        /**
         * Converts the table into a normal range of cells. All data is preserved.
         */
        convertToRange(): Range;

        /**
         * Deletes the table.
         */
        delete(): void;

        /**
         * Gets the range object associated with the data body of the table.
         */
        getDataBodyRange(): Range;

        /**
         * Gets the range object associated with header row of the table.
         */
        getHeaderRowRange(): Range;

        /**
         * Gets the range object associated with the entire table.
         */
        getRange(): Range;

        /**
         * Gets the range object associated with totals row of the table.
         */
        getTotalRowRange(): Range;

        /**
         * Reapplies all the filters currently on the table.
         */
        reapplyFilters(): void;

        getColumns(): TableColumn[];
        addTableColumn(
            index?: number,
            values?:
                | Array<Array<boolean | string | number>>
                | boolean
                | string
                | number,
            name?: string
        ): TableColumn;
        getTableColumn(key: number | string): TableColumn | undefined;

        getRows(): TableRow[];
        addTableRow(
            index?: number,
            values?:
                | Array<Array<boolean | string | number>>
                | boolean
                | string
                | number
        ): TableRow;
    }

    /**
     * Represents a column in a table.
     */
    interface TableColumn {
        /**
         * Retrieve the filter applied to the column. Read-only.
         */
        getFilter(): Filter;

        /**
         * Returns a unique key that identifies the column within the table. Read-only.
         */
        getId(): number;

        /**
         * Returns the index number of the column within the columns collection of the table. Zero-indexed. Read-only.
         */
        getIndex(): number;

        /**
         * Represents the name of the table column.
         */
        getName(): string;
        setName(name: string): void;

        /**
         * Represents the raw values of the specified range. The data returned could be of type string, number, or a boolean. Cells that contain an error will return the error string.
         */
        getValues(): any[][];
        setValues(values: any[][]): void;

        /**
         * Deletes the column from the table.
         */
        delete(): void;

        /**
         * Gets the range object associated with the data body of the column.
         */
        getDataBodyRange(): Range;

        /**
         * Gets the range object associated with the header row of the column.
         */
        getHeaderRowRange(): Range;

        /**
         * Gets the range object associated with the entire column.
         */
        getRange(): Range;

        /**
         * Gets the range object associated with the totals row of the column.
         */
        getTotalRowRange(): Range;
    }

    /**
     * Represents a row in a table.
     *
     *  Note that unlike Ranges or Columns, which will adjust if new rows/columns are added before them,
     *  a TableRow object represent the physical location of the table row, but not the data.
     *  That is, if the data is sorted or if new rows are added, a table row will continue
     *  to point at the index for which it was created.
     */
    interface TableRow {
        /**
         * Represents the raw values of the specified range. The data returned could be of type string, number, or a boolean. Cells that contain an error will return the error string.
         */
        getValues(): any[][];
        setValues(values: any[][]): void;

        /**
         * Deletes the row from the table.
         */
        delete(): void;

        /**
         * Returns the range object associated with the entire row.
         */
        getRange(): Range;
    }

    /**
     * Represents the data validation applied to the current range.
     * To learn more about the data validation object model, read {@link https://docs.microsoft.com/office/dev/add-ins/excel/excel-add-ins-data-validation | Add data validation to Excel ranges}.
     */
    interface DataValidation {
        /**
         * Error alert when user enters invalid data.
         */
        getErrorAlert(): DataValidationErrorAlert;
        setErrorAlert(errorAlert: DataValidationErrorAlert): void;

        /**
         * Ignore blanks: no data validation will be performed on blank cells, it defaults to true.
         */
        getIgnoreBlanks(): boolean;
        setIgnoreBlanks(ignoreBlanks: boolean): void;

        /**
         * Data validation rule that contains different type of data validation criteria.
         */
        getRule(): DataValidationRule;
        setRule(rule: DataValidationRule): void;

        /**
         * Type of the data validation, see Excel.DataValidationType for details.
         */
        getType(): DataValidationType;

        /**
         * Represents if all cell values are valid according to the data validation rules.
         * Returns true if all cell values are valid, or false if all cell values are invalid.
         * Returns null if there are both valid and invalid cell values within the range.
         */
        getValid(): boolean;

        /**
         * Clears the data validation from the current range.
         */
        clear(): void;

        /**
         * Returns a RangeAreas, comprising one or more rectangular ranges, with invalid cell values. If all cell values are valid, this function will return null.
         */
        getInvalidCells(): RangeAreas;
    }

    /**
     * Represents the results from the removeDuplicates method on range
     */
    interface RemoveDuplicatesResult {
        /**
         * Number of duplicated rows removed by the operation.
         */
        getRemoved(): number;

        /**
         * Number of remaining unique rows present in the resulting range.
         */
        getUniqueRemaining(): number;
    }

    /**
     * A format object encapsulating the range's font, fill, borders, alignment, and other properties.
     */
    interface RangeFormat {
        /**
         * Indicates if text is automatically indented when text alignment is set to equal distribution.
         */
        getAutoIndent(): boolean;
        setAutoIndent(autoIndent: boolean): void;

        /**
         * Gets or sets the width of all colums within the range. If the column widths are not uniform, null will be returned.
         */
        getColumnWidth(): number;
        setColumnWidth(columnWidth: number): void;

        /**
         * Returns the fill object defined on the overall range. Read-only.
         */
        getFill(): RangeFill;

        /**
         * Returns the font object defined on the overall range. Read-only.
         */
        getFont(): RangeFont;

        /**
         * Represents the horizontal alignment for the specified object. See Excel.HorizontalAlignment for details.
         */
        getHorizontalAlignment(): HorizontalAlignment;
        setHorizontalAlignment(horizontalAlignment: HorizontalAlignment): void;

        /**
         * An integer from 0 to 250 that indicates the indent level.
         */
        getIndentLevel(): number;
        setIndentLevel(indentLevel: number): void;

        /**
         * Returns the format protection object for a range. Read-only.
         */
        getProtection(): FormatProtection;

        /**
         * The reading order for the range.
         */
        getReadingOrder(): ReadingOrder;
        setReadingOrder(readingOrder: ReadingOrder): void;

        /**
         * Gets or sets the height of all rows in the range. If the row heights are not uniform, null will be returned.
         */
        getRowHeight(): number;
        setRowHeight(rowHeight: number): void;

        /**
         * Indicates if text automatically shrinks to fit in the available column width.
         */
        getShrinkToFit(): boolean;
        setShrinkToFit(shrinkToFit: boolean): void;

        /**
         * Gets or sets the text orientation of all the cells within the range.
         * The text orientation should be an integer either from -90 to 90, or 180 for vertically-oriented text.
         * If the orientation within a range are not uniform, then null will be returned.
         */
        getTextOrientation(): number;
        setTextOrientation(textOrientation: number): void;

        /**
         * Determines if the row height of the Range object equals the standard height of the sheet.
         * Returns True if the row height of the Range object equals the standard height of the sheet.
         * Returns Null if the range contains more than one row and the rows aren't all the same height.
         * Returns False otherwise.
         */
        getUseStandardHeight(): boolean;
        setUseStandardHeight(useStandardHeight: boolean): void;

        /**
         * Indicates whether the column width of the Range object equals the standard width of the sheet.
         * Returns True if the column width of the Range object equals the standard width of the sheet.
         * Returns Null if the range contains more than one column and the columns aren't all the same height.
         * Returns False otherwise.
         */
        getUseStandardWidth(): boolean;
        setUseStandardWidth(useStandardWidth: boolean): void;

        /**
         * Represents the vertical alignment for the specified object. See Excel.VerticalAlignment for details.
         */
        getVerticalAlignment(): VerticalAlignment;
        setVerticalAlignment(verticalAlignment: VerticalAlignment): void;

        /**
         * Indicates if Excel wraps the text in the object. A null value indicates that the entire range doesn't have uniform wrap setting
         */
        getWrapText(): boolean;
        setWrapText(wrapText: boolean): void;

        /**
         * Changes the width of the columns of the current range to achieve the best fit, based on the current data in the columns.
         */
        autofitColumns(): void;

        /**
         * Changes the height of the rows of the current range to achieve the best fit, based on the current data in the columns.
         */
        autofitRows(): void;

        getBorders(): RangeBorder[];
        getRangeBorder(index: BorderIndex): RangeBorder;
    }

    /**
     * Represents the format protection of a range object.
     */
    interface FormatProtection {
        /**
         * Indicates if Excel hides the formula for the cells in the range. A null value indicates that the entire range doesn't have uniform formula hidden setting.
         */
        getFormulaHidden(): boolean;
        setFormulaHidden(formulaHidden: boolean): void;

        /**
         * Indicates if Excel locks the cells in the object. A null value indicates that the entire range doesn't have uniform lock setting.
         */
        getLocked(): boolean;
        setLocked(locked: boolean): void;
    }

    /**
     * Represents the background of a range object.
     */
    interface RangeFill {
        /**
         * HTML color code representing the color of the background, of the form #RRGGBB (e.g. "FFA500") or as a named HTML color (e.g. "orange")
         */
        getColor(): string;
        setColor(color: string): void;

        /**
         * Gets or sets the pattern of a Range. See Excel.FillPattern for details. LinearGradient and RectangularGradient are not supported.
         * A null value indicates that the entire range doesn't have uniform pattern setting.
         */
        getPattern(): FillPattern;
        setPattern(pattern: FillPattern): void;

        /**
         * Sets HTML color code representing the color of the Range pattern, of the form #RRGGBB (e.g. "FFA500") or as a named HTML color (e.g. "orange").
         * Gets HTML color code representing the color of the Range pattern, of the form #RRGGBB (e.g. "FFA500").
         */
        getPatternColor(): string;
        setPatternColor(patternColor: string): void;

        /**
         * Returns or sets a double that lightens or darkens a pattern color for Range Fill, the value is between -1 (darkest) and 1 (brightest), with 0 for the original color.
         * If the pattern tintAndShades are not uniform, null will be returned.
         */
        getPatternTintAndShade(): number;
        setPatternTintAndShade(patternTintAndShade: number): void;

        /**
         * Returns or sets a double that lightens or darkens a color for Range Fill, the value is between -1 (darkest) and 1 (brightest), with 0 for the original color.
         * If the tintAndShades are not uniform, null will be returned.
         */
        getTintAndShade(): number;
        setTintAndShade(tintAndShade: number): void;

        /**
         * Resets the range background.
         */
        clear(): void;
    }

    /**
     * Represents the border of an object.
     */
    interface RangeBorder {
        /**
         * HTML color code representing the color of the border line, of the form #RRGGBB (e.g. "FFA500") or as a named HTML color (e.g. "orange").
         */
        getColor(): string;
        setColor(color: string): void;

        /**
         * Constant value that indicates the specific side of the border. See Excel.BorderIndex for details. Read-only.
         */
        getSideIndex(): BorderIndex;

        /**
         * One of the constants of line style specifying the line style for the border. See Excel.BorderLineStyle for details.
         */
        getStyle(): BorderLineStyle;
        setStyle(style: BorderLineStyle): void;

        /**
         * Returns or sets a double that lightens or darkens a color for Range Border, the value is between -1 (darkest) and 1 (brightest), with 0 for the original color.
         * A null value indicates that the border doesn't have uniform tintAndShade setting.
         */
        getTintAndShade(): number;
        setTintAndShade(tintAndShade: number): void;

        /**
         * Specifies the weight of the border around a range. See Excel.BorderWeight for details.
         */
        getWeight(): BorderWeight;
        setWeight(weight: BorderWeight): void;
    }

    /**
     * This object represents the font attributes (font name, font size, color, etc.) for an object.
     */
    interface RangeFont {
        /**
         * Represents the bold status of font.
         */
        getBold(): boolean;
        setBold(bold: boolean): void;

        /**
         * HTML color code representation of the text color. E.g. #FF0000 represents Red.
         */
        getColor(): string;
        setColor(color: string): void;

        /**
         * Represents the italic status of the font.
         */
        getItalic(): boolean;
        setItalic(italic: boolean): void;

        /**
         * Font name (e.g. "Calibri")
         */
        getName(): string;
        setName(name: string): void;

        /**
         * Font size.
         */
        getSize(): number;
        setSize(size: number): void;

        /**
         * Represents the strikethrough status of font. A null value indicates that the entire range doesn't have uniform Strikethrough setting.
         */
        getStrikethrough(): boolean;
        setStrikethrough(strikethrough: boolean): void;

        /**
         * Represents the Subscript status of font.
         * Returns True if all the fonts of the range are Subscript.
         * Returns False if all the fonts of the range are Superscript or normal (neither Superscript, nor Subscript).
         * Returns Null otherwise.
         */
        getSubscript(): boolean;
        setSubscript(subscript: boolean): void;

        /**
         * Represents the Superscript status of font.
         * Returns True if all the fonts of the range are Superscript.
         * Returns False if all the fonts of the range are Subscript or normal (neither Superscript, nor Subscript).
         * Returns Null otherwise.
         */
        getSuperscript(): boolean;
        setSuperscript(superscript: boolean): void;

        /**
         * Returns or sets a double that lightens or darkens a color for Range Font, the value is between -1 (darkest) and 1 (brightest), with 0 for the original color.
         * A null value indicates that the entire range doesn't have uniform font tintAndShade setting.
         */
        getTintAndShade(): number;
        setTintAndShade(tintAndShade: number): void;

        /**
         * Type of underline applied to the font. See Excel.RangeUnderlineStyle for details.
         */
        getUnderline(): RangeUnderlineStyle;
        setUnderline(underline: RangeUnderlineStyle): void;
    }

    /**
     * Represents a chart object in a workbook.
     * To learn more about the Chart object model, see {@link https://docs.microsoft.com/office/dev/add-ins/excel/excel-add-ins-charts | Work with charts using the Excel JavaScript API}.
     */
    interface Chart {
        /**
         * Represents chart axes. Read-only.
         */
        getAxes(): ChartAxes;

        /**
         * Returns or sets a ChartCategoryLabelLevel enumeration constant referring to
         * the level of where the category labels are being sourced from. Read/Write.
         */
        getCategoryLabelLevel(): number;
        setCategoryLabelLevel(categoryLabelLevel: number): void;

        /**
         * Represents the type of the chart. See Excel.ChartType for details.
         */
        getChartType(): ChartType;
        setChartType(chartType: ChartType): void;

        /**
         * Represents the datalabels on the chart. Read-only.
         */
        getDataLabels(): ChartDataLabels;

        /**
         * Returns or sets the way that blank cells are plotted on a chart. Read/Write.
         */
        getDisplayBlanksAs(): ChartDisplayBlanksAs;
        setDisplayBlanksAs(displayBlanksAs: ChartDisplayBlanksAs): void;

        /**
         * Encapsulates the format properties for the chart area. Read-only.
         */
        getFormat(): ChartAreaFormat;

        /**
         * Represents the height, in points, of the chart object.
         */
        getHeight(): number;
        setHeight(height: number): void;

        /**
         * The unique id of chart. Read-only.
         */
        getId(): string;

        /**
         * The distance, in points, from the left side of the chart to the worksheet origin.
         */
        getLeft(): number;
        setLeft(left: number): void;

        /**
         * Represents the legend for the chart. Read-only.
         */
        getLegend(): ChartLegend;

        /**
         * Represents the name of a chart object.
         */
        getName(): string;
        setName(name: string): void;

        /**
         * Encapsulates the options for a pivot chart. Read-only.
         */
        getPivotOptions(): ChartPivotOptions;

        /**
         * Represents the plotArea for the chart.
         */
        getPlotArea(): ChartPlotArea;

        /**
         * Returns or sets the way columns or rows are used as data series on the chart. Read/Write.
         */
        getPlotBy(): ChartPlotBy;
        setPlotBy(plotBy: ChartPlotBy): void;

        /**
         * True if only visible cells are plotted. False if both visible and hidden cells are plotted. Read/Write.
         */
        getPlotVisibleOnly(): boolean;
        setPlotVisibleOnly(plotVisibleOnly: boolean): void;

        /**
         * Returns or sets a ChartSeriesNameLevel enumeration constant referring to
         * the level of where the series names are being sourced from. Read/Write.
         */
        getSeriesNameLevel(): number;
        setSeriesNameLevel(seriesNameLevel: number): void;

        /**
         * Represents whether to display all field buttons on a PivotChart.
         */
        getShowAllFieldButtons(): boolean;
        setShowAllFieldButtons(showAllFieldButtons: boolean): void;

        /**
         * Represents whether to show the data labels when the value is greater than the maximum value on the value axis.
         * If value axis became smaller than the size of data points, you can use this property to set whether to show the data labels.
         * This property applies to 2-D charts only.
         */
        getShowDataLabelsOverMaximum(): boolean;
        setShowDataLabelsOverMaximum(showDataLabelsOverMaximum: boolean): void;

        /**
         * Returns or sets the chart style for the chart. Read/Write.
         */
        getStyle(): number;
        setStyle(style: number): void;

        /**
         * Represents the title of the specified chart, including the text, visibility, position, and formatting of the title. Read-only.
         */
        getTitle(): ChartTitle;

        /**
         * Represents the distance, in points, from the top edge of the object to the top of row 1 (on a worksheet) or the top of the chart area (on a chart).
         */
        getTop(): number;
        setTop(top: number): void;

        /**
         * Represents the width, in points, of the chart object.
         */
        getWidth(): number;
        setWidth(width: number): void;

        /**
         * The worksheet containing the current chart. Read-only.
         */
        getWorksheet(): Worksheet;

        /**
         * Activates the chart in the Excel UI.
         */
        activate(): void;

        /**
         * Deletes the chart object.
         */
        delete(): void;

        /**
         * Renders the chart as a base64-encoded image by scaling the chart to fit the specified dimensions.
         * The aspect ratio is preserved as part of the resizing.
         */
        getImage(
            width?: number,
            height?: number,
            fittingMode?: ImageFittingMode
        ): string;

        /**
         * Resets the source data for the chart.
         */
        setData(sourceData: Range, seriesBy?: ChartSeriesBy): void;

        /**
         * Positions the chart relative to cells on the worksheet.
         */
        setPosition(startCell: Range | string, endCell?: Range | string): void;

        getSeries(): ChartSeries[];
        addChartSeries(name?: string, index?: number): ChartSeries;
    }

    /**
     * Encapsulates the options for the pivot chart.
     */
    interface ChartPivotOptions {
        /**
         * Specifies whether or not to display the axis field buttons on a PivotChart. The ShowAxisFieldButtons property corresponds to the "Show Axis Field Buttons" command on the "Field Buttons" drop-down list of the "Analyze" tab, which is available when a PivotChart is selected.
         */
        getShowAxisFieldButtons(): boolean;
        setShowAxisFieldButtons(showAxisFieldButtons: boolean): void;

        /**
         * Specifies whether or not to display the legend field buttons on a PivotChart.
         */
        getShowLegendFieldButtons(): boolean;
        setShowLegendFieldButtons(showLegendFieldButtons: boolean): void;

        /**
         * Specifies whether or not to display the report filter field buttons on a PivotChart.
         */
        getShowReportFilterFieldButtons(): boolean;
        setShowReportFilterFieldButtons(
            showReportFilterFieldButtons: boolean
        ): void;

        /**
         * Specifies whether or not to display the show value field buttons on a PivotChart.
         */
        getShowValueFieldButtons(): boolean;
        setShowValueFieldButtons(showValueFieldButtons: boolean): void;
    }

    /**
     * Encapsulates the format properties for the overall chart area.
     */
    interface ChartAreaFormat {
        /**
         * Represents the border format of chart area, which includes color, linestyle, and weight. Read-only.
         */
        getBorder(): ChartBorder;

        /**
         * Returns or sets color scheme of the chart. Read/Write.
         */
        getColorScheme(): ChartColorScheme;
        setColorScheme(colorScheme: ChartColorScheme): void;

        /**
         * Represents the fill format of an object, which includes background formatting information. Read-only.
         */
        getFill(): ChartFill;

        /**
         * Represents the font attributes (font name, font size, color, etc.) for the current object. Read-only.
         */
        getFont(): ChartFont;

        /**
         * Specifies whether or not chart area of the chart has rounded corners. Read/Write.
         */
        getRoundedCorners(): boolean;
        setRoundedCorners(roundedCorners: boolean): void;
    }

    /**
     * Represents a series in a chart.
     */
    interface ChartSeries {
        /**
         * Returns or sets the group for the specified series. Read/Write
         */
        getAxisGroup(): ChartAxisGroup;
        setAxisGroup(axisGroup: ChartAxisGroup): void;

        /**
         * Encapsulates the bin options for histogram charts and pareto charts. Read-only.
         */
        getBinOptions(): ChartBinOptions;

        /**
         * Encapsulates the options for the box and whisker charts. Read-only.
         */
        getBoxwhiskerOptions(): ChartBoxwhiskerOptions;

        /**
         * This can be an integer value from 0 (zero) to 300, representing the percentage of the default size. This property only applies to bubble charts. Read/Write.
         */
        getBubbleScale(): number;
        setBubbleScale(bubbleScale: number): void;

        /**
         * Represents the chart type of a series. See Excel.ChartType for details.
         */
        getChartType(): ChartType;
        setChartType(chartType: ChartType): void;

        /**
         * Represents a collection of all dataLabels in the series.
         */
        getDataLabels(): ChartDataLabels;

        /**
         * Represents the doughnut hole size of a chart series.  Only valid on doughnut and doughnutExploded charts.
         * Throws an invalid argument exception on invalid charts.
         */
        getDoughnutHoleSize(): number;
        setDoughnutHoleSize(doughnutHoleSize: number): void;

        /**
         * Returns or sets the explosion value for a pie-chart or doughnut-chart slice. Returns 0 (zero) if there's no explosion (the tip of the slice is in the center of the pie). Read/Write.
         */
        getExplosion(): number;
        setExplosion(explosion: number): void;

        /**
         * Boolean value representing if the series is filtered or not. Not applicable for surface charts.
         */
        getFiltered(): boolean;
        setFiltered(filtered: boolean): void;

        /**
         * Returns or sets the angle of the first pie-chart or doughnut-chart slice, in degrees (clockwise from vertical). Applies only to pie, 3-D pie, and doughnut charts. Can be a value from 0 through 360. Read/Write
         */
        getFirstSliceAngle(): number;
        setFirstSliceAngle(firstSliceAngle: number): void;

        /**
         * Represents the formatting of a chart series, which includes fill and line formatting. Read-only.
         */
        getFormat(): ChartSeriesFormat;

        /**
         * Represents the gap width of a chart series.  Only valid on bar and column charts, as well as
         * specific classes of line and pie charts.  Throws an invalid argument exception on invalid charts.
         */
        getGapWidth(): number;
        setGapWidth(gapWidth: number): void;

        /**
         * Returns or sets the color for maximum value of a region map chart series. Read/Write.
         */
        getGradientMaximumColor(): string;
        setGradientMaximumColor(gradientMaximumColor: string): void;

        /**
         * Returns or sets the type for maximum value of a region map chart series. Read/Write.
         */
        getGradientMaximumType(): ChartGradientStyleType;
        setGradientMaximumType(
            gradientMaximumType: ChartGradientStyleType
        ): void;

        /**
         * Returns or sets the maximum value of a region map chart series. Read/Write.
         */
        getGradientMaximumValue(): number;
        setGradientMaximumValue(gradientMaximumValue: number): void;

        /**
         * Returns or sets the color for midpoint value of a region map chart series. Read/Write.
         */
        getGradientMidpointColor(): string;
        setGradientMidpointColor(gradientMidpointColor: string): void;

        /**
         * Returns or sets the type for midpoint value of a region map chart series. Read/Write.
         */
        getGradientMidpointType(): ChartGradientStyleType;
        setGradientMidpointType(
            gradientMidpointType: ChartGradientStyleType
        ): void;

        /**
         * Returns or sets the midpoint value of a region map chart series. Read/Write.
         */
        getGradientMidpointValue(): number;
        setGradientMidpointValue(gradientMidpointValue: number): void;

        /**
         * Returns or sets the color for minimum value of a region map chart series. Read/Write.
         */
        getGradientMinimumColor(): string;
        setGradientMinimumColor(gradientMinimumColor: string): void;

        /**
         * Returns or sets the type for minimum value of a region map chart series. Read/Write.
         */
        getGradientMinimumType(): ChartGradientStyleType;
        setGradientMinimumType(
            gradientMinimumType: ChartGradientStyleType
        ): void;

        /**
         * Returns or sets the minimum value of a region map chart series. Read/Write.
         */
        getGradientMinimumValue(): number;
        setGradientMinimumValue(gradientMinimumValue: number): void;

        /**
         * Returns or sets series gradient style of a region map chart. Read/Write.
         */
        getGradientStyle(): ChartGradientStyle;
        setGradientStyle(gradientStyle: ChartGradientStyle): void;

        /**
         * Boolean value representing if the series has data labels or not.
         */
        getHasDataLabels(): boolean;
        setHasDataLabels(hasDataLabels: boolean): void;

        /**
         * Returns or sets the fill color for negative data points in a series. Read/Write.
         */
        getInvertColor(): string;
        setInvertColor(invertColor: string): void;

        /**
         * True if Microsoft Excel inverts the pattern in the item when it corresponds to a negative number. Read/Write.
         */
        getInvertIfNegative(): boolean;
        setInvertIfNegative(invertIfNegative: boolean): void;

        /**
         * Encapsulates the options for a region map chart. Read-only.
         */
        getMapOptions(): ChartMapOptions;

        /**
         * Represents markers background color of a chart series.
         */
        getMarkerBackgroundColor(): string;
        setMarkerBackgroundColor(markerBackgroundColor: string): void;

        /**
         * Represents markers foreground color of a chart series.
         */
        getMarkerForegroundColor(): string;
        setMarkerForegroundColor(markerForegroundColor: string): void;

        /**
         * Represents marker size of a chart series.
         */
        getMarkerSize(): number;
        setMarkerSize(markerSize: number): void;

        /**
         * Represents marker style of a chart series. See Excel.ChartMarkerStyle for details.
         */
        getMarkerStyle(): ChartMarkerStyle;
        setMarkerStyle(markerStyle: ChartMarkerStyle): void;

        /**
         * Represents the name of a series in a chart.
         */
        getName(): string;
        setName(name: string): void;

        /**
         * Specifies how bars and columns are positioned. Can be a value between –100 and 100. Applies only to 2-D bar and 2-D column charts. Read/Write.
         */
        getOverlap(): number;
        setOverlap(overlap: number): void;

        /**
         * Returns or sets the series parent label strategy area for a treemap chart. Read/Write.
         */
        getParentLabelStrategy(): ChartParentLabelStrategy;
        setParentLabelStrategy(
            parentLabelStrategy: ChartParentLabelStrategy
        ): void;

        /**
         * Represents the plot order of a chart series within the chart group.
         */
        getPlotOrder(): number;
        setPlotOrder(plotOrder: number): void;

        /**
         * Returns or sets the size of the secondary section of either a pie of pie chart or a bar of pie chart, as a percentage of the size of the primary pie. Can be a value from 5 to 200. Read/Write.
         */
        getSecondPlotSize(): number;
        setSecondPlotSize(secondPlotSize: number): void;

        /**
         * Specifies whether or not connector lines are shown in waterfall charts. Read/Write.
         */
        getShowConnectorLines(): boolean;
        setShowConnectorLines(showConnectorLines: boolean): void;

        /**
         * Specifies whether or not leader lines are displayed for each data label in the series. Read/Write.
         */
        getShowLeaderLines(): boolean;
        setShowLeaderLines(showLeaderLines: boolean): void;

        /**
         * Boolean value representing if the series has a shadow or not.
         */
        getShowShadow(): boolean;
        setShowShadow(showShadow: boolean): void;

        /**
         * Boolean value representing if the series is smooth or not. Only applicable to line and scatter charts.
         */
        getSmooth(): boolean;
        setSmooth(smooth: boolean): void;

        /**
         * Returns or sets the way the two sections of either a pie of pie chart or a bar of pie chart are split. Read/Write.
         */
        getSplitType(): ChartSplitType;
        setSplitType(splitType: ChartSplitType): void;

        /**
         * Returns or sets the threshold value that separates two sections of either a pie-of-pie chart or a bar-of-pie chart. Read/Write.
         */
        getSplitValue(): number;
        setSplitValue(splitValue: number): void;

        /**
         * True if Microsoft Excel assigns a different color or pattern to each data marker. The chart must contain only one series. Read/Write.
         */
        getVaryByCategories(): boolean;
        setVaryByCategories(varyByCategories: boolean): void;

        /**
         * Represents the error bar object of a chart series.
         */
        getXErrorBars(): ChartErrorBars;

        /**
         * Represents the error bar object of a chart series.
         */
        getYErrorBars(): ChartErrorBars;

        /**
         * Deletes the chart series.
         */
        delete(): void;

        /**
         * Set bubble sizes for a chart series. Only works for bubble charts.
         */
        setBubbleSizes(sourceData: Range): void;

        /**
         * Set values for a chart series. For scatter chart, it means Y axis values.
         */
        setValues(sourceData: Range): void;

        /**
         * Set values of X axis for a chart series. Only works for scatter charts.
         */
        setXAxisValues(sourceData: Range): void;

        getTrendlines(): ChartTrendline[];
        addChartTrendline(type?: ChartTrendlineType): ChartTrendline;
        getChartTrendline(index: number): ChartTrendline;
    }

    /**
     * Encapsulates the format properties for the chart series
     */
    interface ChartSeriesFormat {
        /**
         * Represents the fill format of a chart series, which includes background formatting information. Read-only.
         */
        getFill(): ChartFill;

        /**
         * Represents line formatting. Read-only.
         */
        getLine(): ChartLineFormat;
    }

    /**
     * Represents a point of a series in a chart.
     */
    interface ChartPoint {
        /**
         * Returns the data label of a chart point. Read-only.
         */
        getDataLabel(): ChartDataLabel;

        /**
         * Encapsulates the format properties chart point. Read-only.
         */
        getFormat(): ChartPointFormat;

        /**
         * Represents whether a data point has a data label. Not applicable for surface charts.
         */
        getHasDataLabel(): boolean;
        setHasDataLabel(hasDataLabel: boolean): void;

        /**
         * HTML color code representation of the marker background color of data point. E.g. #FF0000 represents Red.
         */
        getMarkerBackgroundColor(): string;
        setMarkerBackgroundColor(markerBackgroundColor: string): void;

        /**
         * HTML color code representation of the marker foreground color of data point. E.g. #FF0000 represents Red.
         */
        getMarkerForegroundColor(): string;
        setMarkerForegroundColor(markerForegroundColor: string): void;

        /**
         * Represents marker size of data point.
         */
        getMarkerSize(): number;
        setMarkerSize(markerSize: number): void;

        /**
         * Represents marker style of a chart data point. See Excel.ChartMarkerStyle for details.
         */
        getMarkerStyle(): ChartMarkerStyle;
        setMarkerStyle(markerStyle: ChartMarkerStyle): void;
    }

    /**
     * Represents formatting object for chart points.
     */
    interface ChartPointFormat {
        /**
         * Represents the border format of a chart data point, which includes color, style, and weight information. Read-only.
         */
        getBorder(): ChartBorder;

        /**
         * Represents the fill format of a chart, which includes background formatting information. Read-only.
         */
        getFill(): ChartFill;
    }

    /**
     * Represents the chart axes.
     */
    interface ChartAxes {
        /**
         * Represents the category axis in a chart. Read-only.
         */
        getCategoryAxis(): ChartAxis;

        /**
         * Represents the series axis of a 3-dimensional chart. Read-only.
         */
        getSeriesAxis(): ChartAxis;

        /**
         * Represents the value axis in an axis. Read-only.
         */
        getValueAxis(): ChartAxis;
    }

    /**
     * Represents a single axis in a chart.
     */
    interface ChartAxis {
        /**
         * Represents the alignment for the specified axis tick label. See Excel.ChartTextHorizontalAlignment for detail.
         */
        getAlignment(): ChartTickLabelAlignment;
        setAlignment(alignment: ChartTickLabelAlignment): void;

        /**
         * Represents the group for the specified axis. See Excel.ChartAxisGroup for details. Read-only.
         */
        getAxisGroup(): ChartAxisGroup;

        /**
         * Returns or sets the base unit for the specified category axis.
         */
        getBaseTimeUnit(): ChartAxisTimeUnit;
        setBaseTimeUnit(baseTimeUnit: ChartAxisTimeUnit): void;

        /**
         * Returns or sets the category axis type.
         */
        getCategoryType(): ChartAxisCategoryType;
        setCategoryType(categoryType: ChartAxisCategoryType): void;

        /**
         * Represents the custom axis display unit value. Read-only. To set this property, please use the SetCustomDisplayUnit(double) method.
         */
        getCustomDisplayUnit(): number;

        /**
         * Represents the axis display unit. See Excel.ChartAxisDisplayUnit for details.
         */
        getDisplayUnit(): ChartAxisDisplayUnit;
        setDisplayUnit(displayUnit: ChartAxisDisplayUnit): void;

        /**
         * Represents the formatting of a chart object, which includes line and font formatting. Read-only.
         */
        getFormat(): ChartAxisFormat;

        /**
         * Represents the height, in points, of the chart axis. Null if the axis is not visible. Read-only.
         */
        getHeight(): number;

        /**
         * Represents whether value axis crosses the category axis between categories.
         */
        getIsBetweenCategories(): boolean;
        setIsBetweenCategories(isBetweenCategories: boolean): void;

        /**
         * Represents the distance, in points, from the left edge of the axis to the left of chart area. Null if the axis is not visible. Read-only.
         */
        getLeft(): number;

        /**
         * Represents whether or not the number format is linked to the cells. If true, the number format will change in the labels when it changes in the cells.
         */
        getLinkNumberFormat(): boolean;
        setLinkNumberFormat(linkNumberFormat: boolean): void;

        /**
         * Represents the base of the logarithm when using logarithmic scales.
         */
        getLogBase(): number;
        setLogBase(logBase: number): void;

        /**
         * Returns a Gridlines object that represents the major gridlines for the specified axis. Read-only.
         */
        getMajorGridlines(): ChartGridlines;

        /**
         * Represents the type of major tick mark for the specified axis. See Excel.ChartAxisTickMark for details.
         */
        getMajorTickMark(): ChartAxisTickMark;
        setMajorTickMark(majorTickMark: ChartAxisTickMark): void;

        /**
         * Returns or sets the major unit scale value for the category axis when the CategoryType property is set to TimeScale.
         */
        getMajorTimeUnitScale(): ChartAxisTimeUnit;
        setMajorTimeUnitScale(majorTimeUnitScale: ChartAxisTimeUnit): void;

        /**
         * Returns a Gridlines object that represents the minor gridlines for the specified axis. Read-only.
         */
        getMinorGridlines(): ChartGridlines;

        /**
         * Represents the type of minor tick mark for the specified axis. See Excel.ChartAxisTickMark for details.
         */
        getMinorTickMark(): ChartAxisTickMark;
        setMinorTickMark(minorTickMark: ChartAxisTickMark): void;

        /**
         * Returns or sets the minor unit scale value for the category axis when the CategoryType property is set to TimeScale.
         */
        getMinorTimeUnitScale(): ChartAxisTimeUnit;
        setMinorTimeUnitScale(minorTimeUnitScale: ChartAxisTimeUnit): void;

        /**
         * Represents whether an axis is multilevel or not.
         */
        getMultiLevel(): boolean;
        setMultiLevel(multiLevel: boolean): void;

        /**
         * Represents the format code for the axis tick label.
         */
        getNumberFormat(): string;
        setNumberFormat(numberFormat: string): void;

        /**
         * Represents the distance between the levels of labels, and the distance between the first level and the axis line. The value should be an integer from 0 to 1000.
         */
        getOffset(): number;
        setOffset(offset: number): void;

        /**
         * Represents the specified axis position where the other axis crosses. See Excel.ChartAxisPosition for details.
         */
        getPosition(): ChartAxisPosition;
        setPosition(position: ChartAxisPosition): void;

        /**
         * Represents the specified axis position where the other axis crosses at. You should use the SetPositionAt(double) method to set this property.
         */
        getPositionAt(): number;

        /**
         * Represents whether Microsoft Excel plots data points from last to first.
         */
        getReversePlotOrder(): boolean;
        setReversePlotOrder(reversePlotOrder: boolean): void;

        /**
         * Represents the value axis scale type. See Excel.ChartAxisScaleType for details.
         */
        getScaleType(): ChartAxisScaleType;
        setScaleType(scaleType: ChartAxisScaleType): void;

        /**
         * Represents whether the axis display unit label is visible.
         */
        getShowDisplayUnitLabel(): boolean;
        setShowDisplayUnitLabel(showDisplayUnitLabel: boolean): void;

        /**
         * Represents the position of tick-mark labels on the specified axis. See Excel.ChartAxisTickLabelPosition for details.
         */
        getTickLabelPosition(): ChartAxisTickLabelPosition;
        setTickLabelPosition(
            tickLabelPosition: ChartAxisTickLabelPosition
        ): void;

        /**
         * Represents the number of categories or series between tick marks.
         */
        getTickMarkSpacing(): number;
        setTickMarkSpacing(tickMarkSpacing: number): void;

        /**
         * Represents the axis title. Read-only.
         */
        getTitle(): ChartAxisTitle;

        /**
         * Represents the distance, in points, from the top edge of the axis to the top of chart area. Null if the axis is not visible. Read-only.
         */
        getTop(): number;

        /**
         * Represents the axis type. See Excel.ChartAxisType for details.
         */
        getType(): ChartAxisType;

        /**
         * A boolean value represents the visibility of the axis.
         */
        getVisible(): boolean;
        setVisible(visible: boolean): void;

        /**
         * Represents the width, in points, of the chart axis. Null if the axis is not visible. Read-only.
         */
        getWidth(): number;

        /**
         * Sets all the category names for the specified axis.
         */
        setCategoryNames(sourceData: Range): void;

        /**
         * Sets the axis display unit to a custom value.
         */
        setCustomDisplayUnit(value: number): void;

        /**
         * Set the specified axis position where the other axis crosses at.
         */
        setPositionAt(value: number): void;
    }

    /**
     * Encapsulates the format properties for the chart axis.
     */
    interface ChartAxisFormat {
        /**
         * Represents chart fill formatting. Read-only.
         */
        getFill(): ChartFill;

        /**
         * Represents the font attributes (font name, font size, color, etc.) for a chart axis element. Read-only.
         */
        getFont(): ChartFont;

        /**
         * Represents chart line formatting. Read-only.
         */
        getLine(): ChartLineFormat;
    }

    /**
     * Represents the title of a chart axis.
     */
    interface ChartAxisTitle {
        /**
         * Represents the formatting of chart axis title. Read-only.
         */
        getFormat(): ChartAxisTitleFormat;

        /**
         * Represents the axis title.
         */
        getText(): string;
        setText(text: string): void;

        /**
         * A boolean that specifies the visibility of an axis title.
         */
        getVisible(): boolean;
        setVisible(visible: boolean): void;

        /**
         * A string value that represents the formula of chart axis title using A1-style notation.
         */
        setFormula(formula: string): void;
    }

    /**
     * Represents the chart axis title formatting.
     */
    interface ChartAxisTitleFormat {
        /**
         * Represents the border format, which includes color, linestyle, and weight.
         */
        getBorder(): ChartBorder;

        /**
         * Represents chart fill formatting.
         */
        getFill(): ChartFill;

        /**
         * Represents the font attributes, such as font name, font size, color, etc. of chart axis title object. Read-only.
         */
        getFont(): ChartFont;
    }

    /**
     * Represents a collection of all the data labels on a chart point.
     */
    interface ChartDataLabels {
        /**
         * Represents whether data labels automatically generate appropriate text based on context.
         */
        getAutoText(): boolean;
        setAutoText(autoText: boolean): void;

        /**
         * Represents the format of chart data labels, which includes fill and font formatting. Read-only.
         */
        getFormat(): ChartDataLabelFormat;

        /**
         * Represents the horizontal alignment for chart data label. See Excel.ChartTextHorizontalAlignment for details.
         * This property is valid only when TextOrientation of data label is 0.
         */
        getHorizontalAlignment(): ChartTextHorizontalAlignment;
        setHorizontalAlignment(
            horizontalAlignment: ChartTextHorizontalAlignment
        ): void;

        /**
         * Represents whether or not the number format is linked to the cells. If true, the number format will change in the labels when it changes in the cells
         */
        getLinkNumberFormat(): boolean;
        setLinkNumberFormat(linkNumberFormat: boolean): void;

        /**
         * Represents the format code for data labels.
         */
        getNumberFormat(): string;
        setNumberFormat(numberFormat: string): void;

        /**
         * DataLabelPosition value that represents the position of the data label. See Excel.ChartDataLabelPosition for details.
         */
        getPosition(): ChartDataLabelPosition;
        setPosition(position: ChartDataLabelPosition): void;

        /**
         * String representing the separator used for the data labels on a chart.
         */
        getSeparator(): string;
        setSeparator(separator: string): void;

        /**
         * Boolean value representing if the data label bubble size is visible or not.
         */
        getShowBubbleSize(): boolean;
        setShowBubbleSize(showBubbleSize: boolean): void;

        /**
         * Boolean value representing if the data label category name is visible or not.
         */
        getShowCategoryName(): boolean;
        setShowCategoryName(showCategoryName: boolean): void;

        /**
         * Boolean value representing if the data label legend key is visible or not.
         */
        getShowLegendKey(): boolean;
        setShowLegendKey(showLegendKey: boolean): void;

        /**
         * Boolean value representing if the data label percentage is visible or not.
         */
        getShowPercentage(): boolean;
        setShowPercentage(showPercentage: boolean): void;

        /**
         * Boolean value representing if the data label series name is visible or not.
         */
        getShowSeriesName(): boolean;
        setShowSeriesName(showSeriesName: boolean): void;

        /**
         * Boolean value representing if the data label value is visible or not.
         */
        getShowValue(): boolean;
        setShowValue(showValue: boolean): void;

        /**
         * Represents the angle to which the text is oriented for data labels. The value should either be an integer from -90 to 90 or the integer 180 for vertically-oriented text.
         */
        getTextOrientation(): number;
        setTextOrientation(textOrientation: number): void;

        /**
         * Represents the vertical alignment of chart data label. See Excel.ChartTextVerticalAlignment for details.
         * This property is valid only when TextOrientation of data label is -90, 90, or 180.
         */
        getVerticalAlignment(): ChartTextVerticalAlignment;
        setVerticalAlignment(
            verticalAlignment: ChartTextVerticalAlignment
        ): void;
    }

    /**
     * Represents the data label of a chart point.
     */
    interface ChartDataLabel {
        /**
         * Boolean value representing if data label automatically generates appropriate text based on context.
         */
        getAutoText(): boolean;
        setAutoText(autoText: boolean): void;

        /**
         * Represents the format of chart data label.
         */
        getFormat(): ChartDataLabelFormat;

        /**
         * String value that represents the formula of chart data label using A1-style notation.
         */
        getFormula(): string;
        setFormula(formula: string): void;

        /**
         * Returns the height, in points, of the chart data label. Read-only. Null if chart data label is not visible.
         */
        getHeight(): number;

        /**
         * Represents the horizontal alignment for chart data label. See Excel.ChartTextHorizontalAlignment for details.
         * This property is valid only when TextOrientation of data label is -90, 90, or 180.
         */
        getHorizontalAlignment(): ChartTextHorizontalAlignment;
        setHorizontalAlignment(
            horizontalAlignment: ChartTextHorizontalAlignment
        ): void;

        /**
         * Represents the distance, in points, from the left edge of chart data label to the left edge of chart area. Null if chart data label is not visible.
         */
        getLeft(): number;
        setLeft(left: number): void;

        /**
         * Boolean value representing if the number format is linked to the cells (so that the number format changes in the labels when it changes in the cells).
         */
        getLinkNumberFormat(): boolean;
        setLinkNumberFormat(linkNumberFormat: boolean): void;

        /**
         * String value that represents the format code for data label.
         */
        getNumberFormat(): string;
        setNumberFormat(numberFormat: string): void;

        /**
         * DataLabelPosition value that represents the position of the data label. See Excel.ChartDataLabelPosition for details.
         */
        getPosition(): ChartDataLabelPosition;
        setPosition(position: ChartDataLabelPosition): void;

        /**
         * String representing the separator used for the data label on a chart.
         */
        getSeparator(): string;
        setSeparator(separator: string): void;

        /**
         * Boolean value representing if the data label bubble size is visible or not.
         */
        getShowBubbleSize(): boolean;
        setShowBubbleSize(showBubbleSize: boolean): void;

        /**
         * Boolean value representing if the data label category name is visible or not.
         */
        getShowCategoryName(): boolean;
        setShowCategoryName(showCategoryName: boolean): void;

        /**
         * Boolean value representing if the data label legend key is visible or not.
         */
        getShowLegendKey(): boolean;
        setShowLegendKey(showLegendKey: boolean): void;

        /**
         * Boolean value representing if the data label percentage is visible or not.
         */
        getShowPercentage(): boolean;
        setShowPercentage(showPercentage: boolean): void;

        /**
         * Boolean value representing if the data label series name is visible or not.
         */
        getShowSeriesName(): boolean;
        setShowSeriesName(showSeriesName: boolean): void;

        /**
         * Boolean value representing if the data label value is visible or not.
         */
        getShowValue(): boolean;
        setShowValue(showValue: boolean): void;

        /**
         * String representing the text of the data label on a chart.
         */
        getText(): string;
        setText(text: string): void;

        /**
         * Represents the angle to which the text is oriented for the chart data label. The value should either be an integer from -90 to 90 or the integer 180 for vertically-oriented text.
         */
        getTextOrientation(): number;
        setTextOrientation(textOrientation: number): void;

        /**
         * Represents the distance, in points, from the top edge of chart data label to the top of chart area. Null if chart data label is not visible.
         */
        getTop(): number;
        setTop(top: number): void;

        /**
         * Represents the vertical alignment of chart data label. See Excel.ChartTextVerticalAlignment for details.
         * This property is valid only when TextOrientation of data label is 0.
         */
        getVerticalAlignment(): ChartTextVerticalAlignment;
        setVerticalAlignment(
            verticalAlignment: ChartTextVerticalAlignment
        ): void;

        /**
         * Returns the width, in points, of the chart data label. Read-only. Null if chart data label is not visible.
         */
        getWidth(): number;
    }

    /**
     * Encapsulates the format properties for the chart data labels.
     */
    interface ChartDataLabelFormat {
        /**
         * Represents the border format, which includes color, linestyle, and weight. Read-only.
         */
        getBorder(): ChartBorder;

        /**
         * Represents the fill format of the current chart data label. Read-only.
         */
        getFill(): ChartFill;

        /**
         * Represents the font attributes (font name, font size, color, etc.) for a chart data label. Read-only.
         */
        getFont(): ChartFont;
    }

    /**
     * This object represents the attributes for a chart's error bars.
     */
    interface ChartErrorBars {
        /**
         * Specifies whether or not the error bars have an end style cap.
         */
        getEndStyleCap(): boolean;
        setEndStyleCap(endStyleCap: boolean): void;

        /**
         * Specifies the formatting type of the error bars.
         */
        getFormat(): ChartErrorBarsFormat;

        /**
         * Specifies which parts of the error bars to include.
         */
        getInclude(): ChartErrorBarsInclude;
        setInclude(include: ChartErrorBarsInclude): void;

        /**
         * The type of range marked by the error bars.
         */
        getType(): ChartErrorBarsType;
        setType(type: ChartErrorBarsType): void;

        /**
         * Specifies whether or not the error bars are displayed.
         */
        getVisible(): boolean;
        setVisible(visible: boolean): void;
    }

    /**
     * Encapsulates the format properties for chart error bars.
     */
    interface ChartErrorBarsFormat {
        /**
         * Represents the chart line formatting.
         */
        getLine(): ChartLineFormat;
    }

    /**
     * Represents major or minor gridlines on a chart axis.
     */
    interface ChartGridlines {
        /**
         * Represents the formatting of chart gridlines. Read-only.
         */
        getFormat(): ChartGridlinesFormat;

        /**
         * Boolean value representing if the axis gridlines are visible or not.
         */
        getVisible(): boolean;
        setVisible(visible: boolean): void;
    }

    /**
     * Encapsulates the format properties for chart gridlines.
     */
    interface ChartGridlinesFormat {
        /**
         * Represents chart line formatting. Read-only.
         */
        getLine(): ChartLineFormat;
    }

    /**
     * Represents the legend in a chart.
     */
    interface ChartLegend {
        /**
         * Represents the formatting of a chart legend, which includes fill and font formatting. Read-only.
         */
        getFormat(): ChartLegendFormat;

        /**
         * Represents the height, in points, of the legend on the chart. Null if legend is not visible.
         */
        getHeight(): number;
        setHeight(height: number): void;

        /**
         * Represents the left, in points, of a chart legend. Null if legend is not visible.
         */
        getLeft(): number;
        setLeft(left: number): void;

        /**
         * Boolean value for whether the chart legend should overlap with the main body of the chart.
         */
        getOverlay(): boolean;
        setOverlay(overlay: boolean): void;

        /**
         * Represents the position of the legend on the chart. See Excel.ChartLegendPosition for details.
         */
        getPosition(): ChartLegendPosition;
        setPosition(position: ChartLegendPosition): void;

        /**
         * Represents if the legend has a shadow on the chart.
         */
        getShowShadow(): boolean;
        setShowShadow(showShadow: boolean): void;

        /**
         * Represents the top of a chart legend.
         */
        getTop(): number;
        setTop(top: number): void;

        /**
         * A boolean value the represents the visibility of a ChartLegend object.
         */
        getVisible(): boolean;
        setVisible(visible: boolean): void;

        /**
         * Represents the width, in points, of the legend on the chart. Null if legend is not visible.
         */
        getWidth(): number;
        setWidth(width: number): void;

        getLegendEntries(): ChartLegendEntry[];
    }

    /**
     * Represents the legendEntry in legendEntryCollection.
     */
    interface ChartLegendEntry {
        /**
         * Represents the height of the legendEntry on the chart legend.
         */
        getHeight(): number;

        /**
         * Represents the index of the legendEntry in the chart legend.
         */
        getIndex(): number;

        /**
         * Represents the left of a chart legendEntry.
         */
        getLeft(): number;

        /**
         * Represents the top of a chart legendEntry.
         */
        getTop(): number;

        /**
         * Represents the visible of a chart legend entry.
         */
        getVisible(): boolean;
        setVisible(visible: boolean): void;

        /**
         * Represents the width of the legendEntry on the chart Legend.
         */
        getWidth(): number;
    }

    /**
     * Encapsulates the format properties of a chart legend.
     */
    interface ChartLegendFormat {
        /**
         * Represents the border format, which includes color, linestyle, and weight. Read-only.
         */
        getBorder(): ChartBorder;

        /**
         * Represents the fill format of an object, which includes background formatting information. Read-only.
         */
        getFill(): ChartFill;

        /**
         * Represents the font attributes such as font name, font size, color, etc. of a chart legend. Read-only.
         */
        getFont(): ChartFont;
    }

    /**
     * Encapsulates the properties for a region map chart.
     */
    interface ChartMapOptions {
        /**
         * Returns or sets the series map labels strategy of a region map chart. Read/Write.
         */
        getLabelStrategy(): ChartMapLabelStrategy;
        setLabelStrategy(labelStrategy: ChartMapLabelStrategy): void;

        /**
         * Returns or sets the series mapping level of a region map chart. Read/Write.
         */
        getLevel(): ChartMapAreaLevel;
        setLevel(level: ChartMapAreaLevel): void;

        /**
         * Returns or sets the series projection type of a region map chart. Read/Write.
         */
        getProjectionType(): ChartMapProjectionType;
        setProjectionType(projectionType: ChartMapProjectionType): void;
    }

    /**
     * Represents a chart title object of a chart.
     */
    interface ChartTitle {
        /**
         * Represents the formatting of a chart title, which includes fill and font formatting. Read-only.
         */
        getFormat(): ChartTitleFormat;

        /**
         * Returns the height, in points, of the chart title. Null if chart title is not visible. Read-only.
         */
        getHeight(): number;

        /**
         * Represents the horizontal alignment for chart title.
         */
        getHorizontalAlignment(): ChartTextHorizontalAlignment;
        setHorizontalAlignment(
            horizontalAlignment: ChartTextHorizontalAlignment
        ): void;

        /**
         * Represents the distance, in points, from the left edge of chart title to the left edge of chart area. Null if chart title is not visible.
         */
        getLeft(): number;
        setLeft(left: number): void;

        /**
         * Boolean value representing if the chart title will overlay the chart or not.
         */
        getOverlay(): boolean;
        setOverlay(overlay: boolean): void;

        /**
         * Represents the position of chart title. See Excel.ChartTitlePosition for details.
         */
        getPosition(): ChartTitlePosition;
        setPosition(position: ChartTitlePosition): void;

        /**
         * Represents a boolean value that determines if the chart title has a shadow.
         */
        getShowShadow(): boolean;
        setShowShadow(showShadow: boolean): void;

        /**
         * Represents the title text of a chart.
         */
        getText(): string;
        setText(text: string): void;

        /**
         * Represents the angle to which the text is oriented for the chart title. The value should either be an integer from -90 to 90 or the integer 180 for vertically-oriented text.
         */
        getTextOrientation(): number;
        setTextOrientation(textOrientation: number): void;

        /**
         * Represents the distance, in points, from the top edge of chart title to the top of chart area. Null if chart title is not visible.
         */
        getTop(): number;
        setTop(top: number): void;

        /**
         * Represents the vertical alignment of chart title. See Excel.ChartTextVerticalAlignment for details.
         */
        getVerticalAlignment(): ChartTextVerticalAlignment;
        setVerticalAlignment(
            verticalAlignment: ChartTextVerticalAlignment
        ): void;

        /**
         * A boolean value the represents the visibility of a chart title object.
         */
        getVisible(): boolean;
        setVisible(visible: boolean): void;

        /**
         * Returns the width, in points, of the chart title. Null if chart title is not visible. Read-only.
         */
        getWidth(): number;

        /**
         * Get the substring of a chart title. Line break '\n' also counts one character.
         */
        getSubstring(start: number, length: number): ChartFormatString;

        /**
         * Sets a string value that represents the formula of chart title using A1-style notation.
         */
        setFormula(formula: string): void;
    }

    /**
     * Represents the substring in chart related objects that contains text, like ChartTitle object, ChartAxisTitle object, etc.
     */
    interface ChartFormatString {
        /**
         * Represents the font attributes, such as font name, font size, color, etc. of chart characters object.
         */
        getFont(): ChartFont;
    }

    /**
     * Provides access to the office art formatting for chart title.
     */
    interface ChartTitleFormat {
        /**
         * Represents the border format of chart title, which includes color, linestyle, and weight. Read-only.
         */
        getBorder(): ChartBorder;

        /**
         * Represents the fill format of an object, which includes background formatting information. Read-only.
         */
        getFill(): ChartFill;

        /**
         * Represents the font attributes (font name, font size, color, etc.) for an object. Read-only.
         */
        getFont(): ChartFont;
    }

    /**
     * Represents the fill formatting for a chart element.
     */
    interface ChartFill {
        /**
         * Clear the fill color of a chart element.
         */
        clear(): void;

        /**
         * Sets the fill formatting of a chart element to a uniform color.
         */
        setSolidColor(color: string): void;
    }

    /**
     * Represents the border formatting of a chart element.
     */
    interface ChartBorder {
        /**
         * HTML color code representing the color of borders in the chart.
         */
        getColor(): string;
        setColor(color: string): void;

        /**
         * Represents the line style of the border. See Excel.ChartLineStyle for details.
         */
        getLineStyle(): ChartLineStyle;
        setLineStyle(lineStyle: ChartLineStyle): void;

        /**
         * Represents weight of the border, in points.
         */
        getWeight(): number;
        setWeight(weight: number): void;

        /**
         * Clear the border format of a chart element.
         */
        clear(): void;
    }

    /**
     * Encapsulates the bin options for histogram charts and pareto charts.
     */
    interface ChartBinOptions {
        /**
         * Specifies whether or not the bin overflow is enabled in a histogram chart or pareto chart. Read/Write.
         */
        getAllowOverflow(): boolean;
        setAllowOverflow(allowOverflow: boolean): void;

        /**
         * Specifies whether or not the bin underflow is enabled in a histogram chart or pareto chart. Read/Write.
         */
        getAllowUnderflow(): boolean;
        setAllowUnderflow(allowUnderflow: boolean): void;

        /**
         * Returns or sets the bin count of a histogram chart or pareto chart. Read/Write.
         */
        getCount(): number;
        setCount(count: number): void;

        /**
         * Returns or sets the bin overflow value of a histogram chart or pareto chart. Read/Write.
         */
        getOverflowValue(): number;
        setOverflowValue(overflowValue: number): void;

        /**
         * Returns or sets the bin's type for a histogram chart or pareto chart. Read/Write.
         */
        getType(): ChartBinType;
        setType(type: ChartBinType): void;

        /**
         * Returns or sets the bin underflow value of a histogram chart or pareto chart. Read/Write.
         */
        getUnderflowValue(): number;
        setUnderflowValue(underflowValue: number): void;

        /**
         * Returns or sets the bin width value of a histogram chart or pareto chart. Read/Write.
         */
        getWidth(): number;
        setWidth(width: number): void;
    }

    /**
     * Represents the properties of a box and whisker chart.
     */
    interface ChartBoxwhiskerOptions {
        /**
         * Returns or sets the quartile calculation type of a box and whisker chart. Read/Write.
         */
        getQuartileCalculation(): ChartBoxQuartileCalculation;
        setQuartileCalculation(
            quartileCalculation: ChartBoxQuartileCalculation
        ): void;

        /**
         * Specifies whether or not the inner points are shown in a box and whisker chart. Read/Write.
         */
        getShowInnerPoints(): boolean;
        setShowInnerPoints(showInnerPoints: boolean): void;

        /**
         * Specifies whether or not the mean line is shown in a box and whisker chart. Read/Write.
         */
        getShowMeanLine(): boolean;
        setShowMeanLine(showMeanLine: boolean): void;

        /**
         * Specifies whether or not the mean marker is shown in a box and whisker chart. Read/Write.
         */
        getShowMeanMarker(): boolean;
        setShowMeanMarker(showMeanMarker: boolean): void;

        /**
         * Specifies whether or not outlier points are shown in a box and whisker chart. Read/Write.
         */
        getShowOutlierPoints(): boolean;
        setShowOutlierPoints(showOutlierPoints: boolean): void;
    }

    /**
     * Encapsulates the formatting options for line elements.
     */
    interface ChartLineFormat {
        /**
         * HTML color code representing the color of lines in the chart.
         */
        getColor(): string;
        setColor(color: string): void;

        /**
         * Represents the line style. See Excel.ChartLineStyle for details.
         */
        getLineStyle(): ChartLineStyle;
        setLineStyle(lineStyle: ChartLineStyle): void;

        /**
         * Represents weight of the line, in points.
         */
        getWeight(): number;
        setWeight(weight: number): void;

        /**
         * Clear the line format of a chart element.
         */
        clear(): void;
    }

    /**
     * This object represents the font attributes (font name, font size, color, etc.) for a chart object.
     */
    interface ChartFont {
        /**
         * Represents the bold status of font.
         */
        getBold(): boolean;
        setBold(bold: boolean): void;

        /**
         * HTML color code representation of the text color. E.g. #FF0000 represents Red.
         */
        getColor(): string;
        setColor(color: string): void;

        /**
         * Represents the italic status of the font.
         */
        getItalic(): boolean;
        setItalic(italic: boolean): void;

        /**
         * Font name (e.g. "Calibri")
         */
        getName(): string;
        setName(name: string): void;

        /**
         * Size of the font (e.g. 11)
         */
        getSize(): number;
        setSize(size: number): void;

        /**
         * Type of underline applied to the font. See Excel.ChartUnderlineStyle for details.
         */
        getUnderline(): ChartUnderlineStyle;
        setUnderline(underline: ChartUnderlineStyle): void;
    }

    /**
     * This object represents the attributes for a chart trendline object.
     */
    interface ChartTrendline {
        /**
         * Represents the number of periods that the trendline extends backward.
         */
        getBackwardPeriod(): number;
        setBackwardPeriod(backwardPeriod: number): void;

        /**
         * Represents the formatting of a chart trendline.
         */
        getFormat(): ChartTrendlineFormat;

        /**
         * Represents the number of periods that the trendline extends forward.
         */
        getForwardPeriod(): number;
        setForwardPeriod(forwardPeriod: number): void;

        /**
         * Represents the label of a chart trendline.
         */
        getLabel(): ChartTrendlineLabel;

        /**
         * Represents the period of a chart trendline. Only applicable for trendline with MovingAverage type.
         */
        getMovingAveragePeriod(): number;
        setMovingAveragePeriod(movingAveragePeriod: number): void;

        /**
         * Represents the name of the trendline. Can be set to a string value, or can be set to null value represents automatic values. The returned value is always a string
         */
        getName(): string;
        setName(name: string): void;

        /**
         * Represents the order of a chart trendline. Only applicable for trendline with Polynomial type.
         */
        getPolynomialOrder(): number;
        setPolynomialOrder(polynomialOrder: number): void;

        /**
         * True if the equation for the trendline is displayed on the chart.
         */
        getShowEquation(): boolean;
        setShowEquation(showEquation: boolean): void;

        /**
         * True if the R-squared for the trendline is displayed on the chart.
         */
        getShowRSquared(): boolean;
        setShowRSquared(showRSquared: boolean): void;

        /**
         * Represents the type of a chart trendline.
         */
        getType(): ChartTrendlineType;
        setType(type: ChartTrendlineType): void;

        /**
         * Delete the trendline object.
         */
        delete(): void;
    }

    /**
     * Represents the format properties for chart trendline.
     */
    interface ChartTrendlineFormat {
        /**
         * Represents chart line formatting. Read-only.
         */
        getLine(): ChartLineFormat;
    }

    /**
     * This object represents the attributes for a chart trendline lable object.
     */
    interface ChartTrendlineLabel {
        /**
         * Boolean value representing if trendline label automatically generates appropriate text based on context.
         */
        getAutoText(): boolean;
        setAutoText(autoText: boolean): void;

        /**
         * Represents the format of chart trendline label.
         */
        getFormat(): ChartTrendlineLabelFormat;

        /**
         * String value that represents the formula of chart trendline label using A1-style notation.
         */
        getFormula(): string;
        setFormula(formula: string): void;

        /**
         * Returns the height, in points, of the chart trendline label. Read-only. Null if chart trendline label is not visible.
         */
        getHeight(): number;

        /**
         * Represents the horizontal alignment for chart trendline label. See Excel.ChartTextHorizontalAlignment for details.
         * This property is valid only when TextOrientation of trendline label is -90, 90, or 180.
         */
        getHorizontalAlignment(): ChartTextHorizontalAlignment;
        setHorizontalAlignment(
            horizontalAlignment: ChartTextHorizontalAlignment
        ): void;

        /**
         * Represents the distance, in points, from the left edge of chart trendline label to the left edge of chart area. Null if chart trendline label is not visible.
         */
        getLeft(): number;
        setLeft(left: number): void;

        /**
         * Boolean value representing if the number format is linked to the cells (so that the number format changes in the labels when it changes in the cells).
         */
        getLinkNumberFormat(): boolean;
        setLinkNumberFormat(linkNumberFormat: boolean): void;

        /**
         * String value that represents the format code for trendline label.
         */
        getNumberFormat(): string;
        setNumberFormat(numberFormat: string): void;

        /**
         * String representing the text of the trendline label on a chart.
         */
        getText(): string;
        setText(text: string): void;

        /**
         * Represents the angle to which the text is oriented for the chart trendline label. The value should either be an integer from -90 to 90 or the integer 180 for vertically-oriented text.
         */
        getTextOrientation(): number;
        setTextOrientation(textOrientation: number): void;

        /**
         * Represents the distance, in points, from the top edge of chart trendline label to the top of chart area. Null if chart trendline label is not visible.
         */
        getTop(): number;
        setTop(top: number): void;

        /**
         * Represents the vertical alignment of chart trendline label. See Excel.ChartTextVerticalAlignment for details.
         * This property is valid only when TextOrientation of trendline label is 0.
         */
        getVerticalAlignment(): ChartTextVerticalAlignment;
        setVerticalAlignment(
            verticalAlignment: ChartTextVerticalAlignment
        ): void;

        /**
         * Returns the width, in points, of the chart trendline label. Read-only. Null if chart trendline label is not visible.
         */
        getWidth(): number;
    }

    /**
     * Encapsulates the format properties for the chart trendline label.
     */
    interface ChartTrendlineLabelFormat {
        /**
         * Represents the border format, which includes color, linestyle, and weight.
         */
        getBorder(): ChartBorder;

        /**
         * Represents the fill format of the current chart trendline label.
         */
        getFill(): ChartFill;

        /**
         * Represents the font attributes (font name, font size, color, etc.) for a chart trendline label.
         */
        getFont(): ChartFont;
    }

    /**
     * This object represents the attributes for a chart plotArea object.
     */
    interface ChartPlotArea {
        /**
         * Represents the formatting of a chart plotArea.
         */
        getFormat(): ChartPlotAreaFormat;

        /**
         * Represents the height value of plotArea.
         */
        getHeight(): number;
        setHeight(height: number): void;

        /**
         * Represents the insideHeight value of plotArea.
         */
        getInsideHeight(): number;
        setInsideHeight(insideHeight: number): void;

        /**
         * Represents the insideLeft value of plotArea.
         */
        getInsideLeft(): number;
        setInsideLeft(insideLeft: number): void;

        /**
         * Represents the insideTop value of plotArea.
         */
        getInsideTop(): number;
        setInsideTop(insideTop: number): void;

        /**
         * Represents the insideWidth value of plotArea.
         */
        getInsideWidth(): number;
        setInsideWidth(insideWidth: number): void;

        /**
         * Represents the left value of plotArea.
         */
        getLeft(): number;
        setLeft(left: number): void;

        /**
         * Represents the position of plotArea.
         */
        getPosition(): ChartPlotAreaPosition;
        setPosition(position: ChartPlotAreaPosition): void;

        /**
         * Represents the top value of plotArea.
         */
        getTop(): number;
        setTop(top: number): void;

        /**
         * Represents the width value of plotArea.
         */
        getWidth(): number;
        setWidth(width: number): void;
    }

    /**
     * Represents the format properties for chart plotArea.
     */
    interface ChartPlotAreaFormat {
        /**
         * Represents the border attributes of a chart plotArea.
         */
        getBorder(): ChartBorder;

        /**
         * Represents the fill format of an object, which includes background formatting information.
         */
        getFill(): ChartFill;
    }

    /**
     * Manages sorting operations on Range objects.
     */
    interface RangeSort {
        /**
         * Perform a sort operation.
         */
        apply(
            fields: SortField[],
            matchCase?: boolean,
            hasHeaders?: boolean,
            orientation?: SortOrientation,
            method?: SortMethod
        ): void;
    }

    /**
     * Manages sorting operations on Table objects.
     */
    interface TableSort {
        /**
         * Represents whether the casing impacted the last sort of the table. Read-only.
         */
        getMatchCase(): boolean;

        /**
         * Represents Chinese character ordering method last used to sort the table. Read-only.
         */
        getMethod(): SortMethod;

        /**
         * Perform a sort operation.
         */
        apply(
            fields: SortField[],
            matchCase?: boolean,
            method?: SortMethod
        ): void;

        /**
         * Clears the sorting that is currently on the table. While this doesn't modify the table's ordering, it clears the state of the header buttons.
         */
        clear(): void;

        /**
         * Reapplies the current sorting parameters to the table.
         */
        reapply(): void;
    }

    /**
     * Manages the filtering of a table's column.
     */
    interface Filter {
        /**
         * The currently applied filter on the given column. Read-only.
         */
        getCriteria(): FilterCriteria;

        /**
         * Apply the given filter criteria on the given column.
         */
        apply(criteria: FilterCriteria): void;

        /**
         * Apply a "Bottom Item" filter to the column for the given number of elements.
         */
        applyBottomItemsFilter(count: number): void;

        /**
         * Apply a "Bottom Percent" filter to the column for the given percentage of elements.
         */
        applyBottomPercentFilter(percent: number): void;

        /**
         * Apply a "Cell Color" filter to the column for the given color.
         */
        applyCellColorFilter(color: string): void;

        /**
         * Apply an "Icon" filter to the column for the given criteria strings.
         */
        applyCustomFilter(
            criteria1: string,
            criteria2?: string,
            oper?: FilterOperator
        ): void;

        /**
         * Apply a "Dynamic" filter to the column.
         */
        applyDynamicFilter(criteria: DynamicFilterCriteria): void;

        /**
         * Apply a "Font Color" filter to the column for the given color.
         */
        applyFontColorFilter(color: string): void;

        /**
         * Apply an "Icon" filter to the column for the given icon.
         */
        applyIconFilter(icon: Icon): void;

        /**
         * Apply a "Top Item" filter to the column for the given number of elements.
         */
        applyTopItemsFilter(count: number): void;

        /**
         * Apply a "Top Percent" filter to the column for the given percentage of elements.
         */
        applyTopPercentFilter(percent: number): void;

        /**
         * Apply a "Values" filter to the column for the given values.
         */
        applyValuesFilter(values: Array<string | FilterDatetime>): void;

        /**
         * Clear the filter on the given column.
         */
        clear(): void;
    }

    /**
     * Represents the AutoFilter object.
     *  AutoFilter turns the values in Excel column into specific filters based on the cell contents.
     */
    interface AutoFilter {
        /**
         * Indicates if the AutoFilter is enabled or not. Read-Only.
         */
        getEnabled(): boolean;

        /**
         * Indicates if the AutoFilter has filter criteria. Read-Only.
         */
        getIsDataFiltered(): boolean;

        /**
         * Applies the AutoFilter to a range. This filters the column if column index and filter criteria are specified.
         */
        apply(
            range: Range | string,
            columnIndex?: number,
            criteria?: FilterCriteria
        ): void;

        /**
         * Clears the filter criteria of the AutoFilter.
         */
        clearCriteria(): void;

        /**
         * Returns the Range object that represents the range to which the AutoFilter applies.
         * If there is no Range object associated with the AutoFilter, this method returns a null object.
         */
        getRange(): Range;

        /**
         * Applies the specified Autofilter object currently on the range.
         */
        reapply(): void;

        /**
         * Removes the AutoFilter for the range.
         */
        remove(): void;
    }

    /**
     * Represents a custom XML part object in a workbook.
     */
    interface CustomXmlPart {
        /**
         * The custom XML part's ID. Read-only.
         */
        getId(): string;

        /**
         * The custom XML part's namespace URI. Read-only.
         */
        getNamespaceUri(): string;

        /**
         * Deletes the custom XML part.
         */
        delete(): void;

        /**
         * Gets the custom XML part's full XML content.
         */
        getXml(): string;

        /**
         * Sets the custom XML part's full XML content.
         */
        setXml(xml: string): void;
    }

    /**
     * Represents an Excel PivotTable.
     * To learn more about the PivotTable object model, read {@link https://docs.microsoft.com/office/dev/add-ins/excel/excel-add-ins-pivottables | Work with PivotTables using the Excel JavaScript API}.
     */
    interface PivotTable {
        /**
         * Specifies whether the PivotTable allows values in the data body to be edited by the user.
         */
        getEnableDataValueEditing(): boolean;
        setEnableDataValueEditing(enableDataValueEditing: boolean): void;

        /**
         * Id of the PivotTable. Read-only.
         */
        getId(): string;

        /**
         * The PivotLayout describing the layout and visual structure of the PivotTable.
         */
        getLayout(): PivotLayout;

        /**
         * Name of the PivotTable.
         */
        getName(): string;
        setName(name: string): void;

        /**
         * Specifies whether the PivotTable uses custom lists when sorting.
         */
        getUseCustomSortLists(): boolean;
        setUseCustomSortLists(useCustomSortLists: boolean): void;

        /**
         * The worksheet containing the current PivotTable.
         */
        getWorksheet(): Worksheet;

        /**
         * Deletes the PivotTable.
         */
        delete(): void;

        /**
         * Refreshes the PivotTable.
         */
        refresh(): void;

        getDataHierarchies(): DataPivotHierarchy[];
        addDataPivotHierarchy(
            pivotHierarchy: PivotHierarchy
        ): DataPivotHierarchy;
        getDataPivotHierarchy(name: string): DataPivotHierarchy | undefined;
        removeDataPivotHierarchy(DataPivotHierarchy: DataPivotHierarchy): void;

        getFilterHierarchies(): FilterPivotHierarchy[];
        addFilterPivotHierarchy(
            pivotHierarchy: PivotHierarchy
        ): FilterPivotHierarchy;
        getFilterPivotHierarchy(name: string): FilterPivotHierarchy | undefined;
        removeFilterPivotHierarchy(
            filterPivotHierarchy: FilterPivotHierarchy
        ): void;

        getHierarchies(): PivotHierarchy[];
        getPivotHierarchy(name: string): PivotHierarchy | undefined;
    }

    /**
     * Represents the visual layout of the PivotTable.
     */
    interface PivotLayout {
        /**
         * Specifies whether formatting will be automatically formatted when it’s refreshed or when fields are moved
         */
        getAutoFormat(): boolean;
        setAutoFormat(autoFormat: boolean): void;

        /**
         * Specifies whether the field list can be shown in the UI.
         */
        getEnableFieldList(): boolean;
        setEnableFieldList(enableFieldList: boolean): void;

        /**
         * This property indicates the PivotLayoutType of all fields on the PivotTable. If fields have different states, this will be null.
         */
        getLayoutType(): PivotLayoutType;
        setLayoutType(layoutType: PivotLayoutType): void;

        /**
         * Specifies whether formatting is preserved when the report is refreshed or recalculated by operations such as pivoting, sorting, or changing page field items.
         */
        getPreserveFormatting(): boolean;
        setPreserveFormatting(preserveFormatting: boolean): void;

        /**
         * Specifies whether the PivotTable report shows grand totals for columns.
         */
        getShowColumnGrandTotals(): boolean;
        setShowColumnGrandTotals(showColumnGrandTotals: boolean): void;

        /**
         * Specifies whether the PivotTable report shows grand totals for rows.
         */
        getShowRowGrandTotals(): boolean;
        setShowRowGrandTotals(showRowGrandTotals: boolean): void;

        /**
         * This property indicates the SubtotalLocationType of all fields on the PivotTable. If fields have different states, this will be null.
         */
        getSubtotalLocation(): SubtotalLocationType;
        setSubtotalLocation(subtotalLocation: SubtotalLocationType): void;

        /**
         * Returns the range where the PivotTable's column labels reside.
         */
        getColumnLabelRange(): Range;

        /**
         * Returns the range where the PivotTable's data values reside.
         */
        getDataBodyRange(): Range;

        /**
         * Gets the DataHierarchy that is used to calculate the value in a specified range within the PivotTable.
         */
        getDataHierarchy(cell: Range | string): DataPivotHierarchy;

        /**
         * Returns the range of the PivotTable's filter area.
         */
        getFilterAxisRange(): Range;

        /**
         * Returns the range the PivotTable exists on, excluding the filter area.
         */
        getRange(): Range;

        /**
         * Returns the range where the PivotTable's row labels reside.
         */
        getRowLabelRange(): Range;

        /**
         * Sets the PivotTable to automatically sort using the specified cell to automatically select all necessary criteria and context. This behaves identically to applying an autosort from the UI.
         */
        setAutoSortOnCell(cell: Range | string, sortBy: SortBy): void;
    }

    /**
     * Represents the Excel PivotHierarchy.
     */
    interface PivotHierarchy {
        /**
         * Id of the PivotHierarchy.
         */
        getId(): string;

        /**
         * Name of the PivotHierarchy.
         */
        getName(): string;
        setName(name: string): void;

        getFields(): PivotField[];
        getPivotField(name: string): PivotField | undefined;
    }

    /**
     * Represents the Excel RowColumnPivotHierarchy.
     */
    interface RowColumnPivotHierarchy {
        /**
         * Id of the RowColumnPivotHierarchy.
         */
        getId(): string;

        /**
         * Name of the RowColumnPivotHierarchy.
         */
        getName(): string;
        setName(name: string): void;

        /**
         * Position of the RowColumnPivotHierarchy.
         */
        getPosition(): number;
        setPosition(position: number): void;

        /**
         * Reset the RowColumnPivotHierarchy back to its default values.
         */
        setToDefault(): void;

        getFields(): PivotField[];
        getPivotField(name: string): PivotField | undefined;
    }

    /**
     * Represents the Excel FilterPivotHierarchy.
     */
    interface FilterPivotHierarchy {
        /**
         * Determines whether to allow multiple filter items.
         */
        getEnableMultipleFilterItems(): boolean;
        setEnableMultipleFilterItems(enableMultipleFilterItems: boolean): void;

        /**
         * Id of the FilterPivotHierarchy.
         */
        getId(): string;

        /**
         * Name of the FilterPivotHierarchy.
         */
        getName(): string;
        setName(name: string): void;

        /**
         * Position of the FilterPivotHierarchy.
         */
        getPosition(): number;
        setPosition(position: number): void;

        /**
         * Reset the FilterPivotHierarchy back to its default values.
         */
        setToDefault(): void;

        getFields(): PivotField[];
        getPivotField(name: string): PivotField | undefined;
    }

    /**
     * Represents the Excel DataPivotHierarchy.
     */
    interface DataPivotHierarchy {
        /**
         * Returns the PivotFields associated with the DataPivotHierarchy.
         */
        getField(): PivotField;

        /**
         * Id of the DataPivotHierarchy.
         */
        getId(): string;

        /**
         * Name of the DataPivotHierarchy.
         */
        getName(): string;
        setName(name: string): void;

        /**
         * Number format of the DataPivotHierarchy.
         */
        getNumberFormat(): string;
        setNumberFormat(numberFormat: string): void;

        /**
         * Position of the DataPivotHierarchy.
         */
        getPosition(): number;
        setPosition(position: number): void;

        /**
         * Determines whether the data should be shown as a specific summary calculation or not.
         */
        getShowAs(): ShowAsRule;
        setShowAs(showAs: ShowAsRule): void;

        /**
         * Determines whether to show all items of the DataPivotHierarchy.
         */
        getSummarizeBy(): AggregationFunction;
        setSummarizeBy(summarizeBy: AggregationFunction): void;

        /**
         * Reset the DataPivotHierarchy back to its default values.
         */
        setToDefault(): void;
    }

    /**
     * Represents the Excel PivotField.
     */
    interface PivotField {
        /**
         * Id of the PivotField.
         */
        getId(): string;

        /**
         * Name of the PivotField.
         */
        getName(): string;
        setName(name: string): void;

        /**
         * Determines whether to show all items of the PivotField.
         */
        getShowAllItems(): boolean;
        setShowAllItems(showAllItems: boolean): void;

        /**
         * Subtotals of the PivotField.
         */
        getSubtotals(): Subtotals;
        setSubtotals(subtotals: Subtotals): void;

        /**
         * Sorts the PivotField. If a DataPivotHierarchy is specified, then sort will be applied based on it, if not sort will be based on the PivotField itself.
         */
        sortByLabels(sortBy: SortBy): void;

        /**
         * Sorts the PivotField by specified values in a given scope. The scope defines which specific values will be used to sort when
         * there are multiple values from the same DataPivotHierarchy.
         */
        sortByValues(
            sortBy: SortBy,
            valuesHierarchy: DataPivotHierarchy,
            pivotItemScope?: Array<PivotItem | string>
        ): void;

        getItems(): PivotItem[];
        getPivotItem(name: string): PivotItem | undefined;
    }

    /**
     * Represents the Excel PivotItem.
     */
    interface PivotItem {
        /**
         * Id of the PivotItem.
         */
        getId(): string;

        /**
         * Determines whether the item is expanded to show child items or if it's collapsed and child items are hidden.
         */
        getIsExpanded(): boolean;
        setIsExpanded(isExpanded: boolean): void;

        /**
         * Name of the PivotItem.
         */
        getName(): string;
        setName(name: string): void;

        /**
         * Determines whether the PivotItem is visible or not.
         */
        getVisible(): boolean;
        setVisible(visible: boolean): void;
    }

    /**
     * Represents workbook properties.
     */
    interface DocumentProperties {
        /**
         * Gets or sets the author of the workbook.
         */
        getAuthor(): string;
        setAuthor(author: string): void;

        /**
         * Gets or sets the category of the workbook.
         */
        getCategory(): string;
        setCategory(category: string): void;

        /**
         * Gets or sets the comments of the workbook.
         */
        getComments(): string;
        setComments(comments: string): void;

        /**
         * Gets or sets the company of the workbook.
         */
        getCompany(): string;
        setCompany(company: string): void;

        /**
         * Gets the creation date of the workbook. Read only.
         */
        getCreationDate(): Date;

        /**
         * Gets or sets the keywords of the workbook.
         */
        getKeywords(): string;
        setKeywords(keywords: string): void;

        /**
         * Gets the last author of the workbook. Read only.
         */
        getLastAuthor(): string;

        /**
         * Gets or sets the manager of the workbook.
         */
        getManager(): string;
        setManager(manager: string): void;

        /**
         * Gets the revision number of the workbook. Read only.
         */
        getRevisionNumber(): number;
        setRevisionNumber(revisionNumber: number): void;

        /**
         * Gets or sets the subject of the workbook.
         */
        getSubject(): string;
        setSubject(subject: string): void;

        /**
         * Gets or sets the title of the workbook.
         */
        getTitle(): string;
        setTitle(title: string): void;

        getCustom(): CustomProperty[];
        addCustomProperty(key: string, value: any): CustomProperty;
        deleteAllCustomProperties(): void;
        getCustomProperty(key: string): CustomProperty | undefined;
    }

    /**
     * Represents a custom property.
     */
    interface CustomProperty {
        /**
         * Gets the key of the custom property. Read only.
         */
        getKey(): string;

        /**
         * Gets the value type of the custom property. Read only.
         */
        getType(): DocumentPropertyType;

        /**
         * Gets or sets the value of the custom property.
         */
        getValue(): any;
        setValue(value: any): void;

        /**
         * Deletes the custom property.
         */
        delete(): void;
    }

    /**
     * An object encapsulating a conditional format's range, format, rule, and other properties.
     * To learn more about the conditional formatting object model, read {@link https://docs.microsoft.com/office/dev/add-ins/excel/excel-add-ins-conditional-formatting | Apply conditional formatting to Excel ranges}.
     */
    interface ConditionalFormat {
        /**
         * Returns the cell value conditional format properties if the current conditional format is a CellValue type.
         * For example to format all cells between 5 and 10. Read-only.
         */
        getCellValue(): CellValueConditionalFormat | undefined;

        /**
         * Returns the ColorScale conditional format properties if the current conditional format is an ColorScale type. Read-only.
         */
        getColorScale(): ColorScaleConditionalFormat | undefined;

        /**
         * Returns the custom conditional format properties if the current conditional format is a custom type. Read-only.
         */
        getCustom(): CustomConditionalFormat | undefined;

        /**
         * Returns the data bar properties if the current conditional format is a data bar. Read-only.
         */
        getDataBar(): DataBarConditionalFormat | undefined;

        /**
         * Returns the IconSet conditional format properties if the current conditional format is an IconSet type. Read-only.
         */
        getIconSet(): IconSetConditionalFormat | undefined;

        /**
         * The Priority of the Conditional Format within the current ConditionalFormatCollection. Read-only.
         */
        getId(): string;

        /**
         * Returns the preset criteria conditional format. See Excel.PresetCriteriaConditionalFormat for more details.
         */
        getPreset(): PresetCriteriaConditionalFormat | undefined;

        /**
         * The priority (or index) within the conditional format collection that this conditional format currently exists in. Changing this also
         * changes other conditional formats' priorities, to allow for a contiguous priority order.
         * Use a negative priority to begin from the back.
         * Priorities greater than than bounds will get and set to the maximum (or minimum if negative) priority.
         * Also note that if you change the priority, you have to re-fetch a new copy of the object at that new priority location if you want to make further changes to it. Read-only.
         */
        getPriority(): number;
        setPriority(priority: number): void;

        /**
         * If the conditions of this conditional format are met, no lower-priority formats shall take effect on that cell.
         * Null on databars, icon sets, and colorscales as there's no concept of StopIfTrue for these
         */
        getStopIfTrue(): boolean;
        setStopIfTrue(stopIfTrue: boolean): void;

        /**
         * Returns the specific text conditional format properties if the current conditional format is a text type.
         * For example to format cells matching the word "Text". Read-only.
         */
        getTextComparison(): TextConditionalFormat | undefined;

        /**
         * Returns the Top/Bottom conditional format properties if the current conditional format is an TopBottom type.
         * For example to format the top 10% or bottom 10 items. Read-only.
         */
        getTopBottom(): TopBottomConditionalFormat | undefined;

        /**
         * A type of conditional format. Only one can be set at a time. Read-only.
         */
        getType(): ConditionalFormatType;

        /**
         * Deletes this conditional format.
         */
        delete(): void;

        /**
         * Returns the range the conditonal format is applied to, or a null object if the conditional format is applied to multiple ranges. Read-only.
         */
        getRange(): Range;

        /**
         * Returns the RangeAreas, comprising one or more rectangular ranges, the conditonal format is applied to. Read-only.
         */
        getRanges(): RangeAreas;
    }

    /**
     * Represents an Excel Conditional Data Bar Type.
     */
    interface DataBarConditionalFormat {
        /**
         * HTML color code representing the color of the Axis line, of the form #RRGGBB (e.g. "FFA500") or as a named HTML color (e.g. "orange").
         * "" (empty string) if no axis is present or set.
         */
        getAxisColor(): string;
        setAxisColor(axisColor: string): void;

        /**
         * Representation of how the axis is determined for an Excel data bar.
         */
        getAxisFormat(): ConditionalDataBarAxisFormat;
        setAxisFormat(axisFormat: ConditionalDataBarAxisFormat): void;

        /**
         * Represents the direction that the data bar graphic should be based on.
         */
        getBarDirection(): ConditionalDataBarDirection;
        setBarDirection(barDirection: ConditionalDataBarDirection): void;

        /**
         * The rule for what consistutes the lower bound (and how to calculate it, if applicable) for a data bar.
         * The `ConditionalDataBarRule` object must be set as a JSON object (use `x.lowerBoundRule = {...}` instead of `x.lowerBoundRule.formula = ...`).
         */
        getLowerBoundRule(): ConditionalDataBarRule;
        setLowerBoundRule(lowerBoundRule: ConditionalDataBarRule): void;

        /**
         * Representation of all values to the left of the axis in an Excel data bar. Read-only.
         */
        getNegativeFormat(): ConditionalDataBarNegativeFormat;

        /**
         * Representation of all values to the right of the axis in an Excel data bar. Read-only.
         */
        getPositiveFormat(): ConditionalDataBarPositiveFormat;

        /**
         * If true, hides the values from the cells where the data bar is applied.
         */
        getShowDataBarOnly(): boolean;
        setShowDataBarOnly(showDataBarOnly: boolean): void;

        /**
         * The rule for what constitutes the upper bound (and how to calculate it, if applicable) for a data bar.
         * The `ConditionalDataBarRule` object must be set as a JSON object (use `x.upperBoundRule = {...}` instead of `x.upperBoundRule.formula = ...`).
         */
        getUpperBoundRule(): ConditionalDataBarRule;
        setUpperBoundRule(upperBoundRule: ConditionalDataBarRule): void;
    }

    /**
     * Represents a conditional format DataBar Format for the positive side of the data bar.
     */
    interface ConditionalDataBarPositiveFormat {
        /**
         * HTML color code representing the color of the border line, of the form #RRGGBB (e.g. "FFA500") or as a named HTML color (e.g. "orange").
         * "" (empty string) if no border is present or set.
         */
        getBorderColor(): string;
        setBorderColor(borderColor: string): void;

        /**
         * HTML color code representing the fill color, of the form #RRGGBB (e.g. "FFA500") or as a named HTML color (e.g. "orange").
         */
        getFillColor(): string;
        setFillColor(fillColor: string): void;

        /**
         * Boolean representation of whether or not the DataBar has a gradient.
         */
        getGradientFill(): boolean;
        setGradientFill(gradientFill: boolean): void;
    }

    /**
     * Represents a conditional format DataBar Format for the negative side of the data bar.
     */
    interface ConditionalDataBarNegativeFormat {
        /**
         * HTML color code representing the color of the border line, of the form #RRGGBB (e.g. "FFA500") or as a named HTML color (e.g. "orange").
         * "Empty String" if no border is present or set.
         */
        getBorderColor(): string;
        setBorderColor(borderColor: string): void;

        /**
         * HTML color code representing the fill color, of the form #RRGGBB (e.g. "FFA500") or as a named HTML color (e.g. "orange").
         */
        getFillColor(): string;
        setFillColor(fillColor: string): void;

        /**
         * Boolean representation of whether or not the negative DataBar has the same border color as the positive DataBar.
         */
        getMatchPositiveBorderColor(): boolean;
        setMatchPositiveBorderColor(matchPositiveBorderColor: boolean): void;

        /**
         * Boolean representation of whether or not the negative DataBar has the same fill color as the positive DataBar.
         */
        getMatchPositiveFillColor(): boolean;
        setMatchPositiveFillColor(matchPositiveFillColor: boolean): void;
    }

    /**
     * Represents a custom conditional format type.
     */
    interface CustomConditionalFormat {
        /**
         * Returns a format object, encapsulating the conditional formats font, fill, borders, and other properties. Read-only.
         */
        getFormat(): ConditionalRangeFormat;

        /**
         * Represents the Rule object on this conditional format. Read-only.
         */
        getRule(): ConditionalFormatRule;
    }

    /**
     * Represents a rule, for all traditional rule/format pairings.
     */
    interface ConditionalFormatRule {
        /**
         * The formula, if required, to evaluate the conditional format rule on.
         */
        getFormula(): string;
        setFormula(formula: string): void;

        /**
         * The formula, if required, to evaluate the conditional format rule on in the user's language.
         */
        getFormulaLocal(): string;
        setFormulaLocal(formulaLocal: string): void;
    }

    /**
     * Represents an IconSet criteria for conditional formatting.
     */
    interface IconSetConditionalFormat {
        /**
         * If true, reverses the icon orders for the IconSet. Note that this cannot be set if custom icons are used.
         */
        getReverseIconOrder(): boolean;
        setReverseIconOrder(reverseIconOrder: boolean): void;

        /**
         * If true, hides the values and only shows icons.
         */
        getShowIconOnly(): boolean;
        setShowIconOnly(showIconOnly: boolean): void;

        /**
         * If set, displays the IconSet option for the conditional format.
         */
        getStyle(): IconSet;
        setStyle(style: IconSet): void;
    }

    /**
     * Represents an IconSet criteria for conditional formatting.
     */
    interface ColorScaleConditionalFormat {
        /**
         * The criteria of the color scale. Midpoint is optional when using a two point color scale.
         */
        getCriteria(): ConditionalColorScaleCriteria;
        setCriteria(criteria: ConditionalColorScaleCriteria): void;

        /**
         * If true the color scale will have three points (minimum, midpoint, maximum), otherwise it will have two (minimum, maximum).
         */
        getThreeColorScale(): boolean;
    }

    /**
     * Represents a Top/Bottom conditional format.
     */
    interface TopBottomConditionalFormat {
        /**
         * Returns a format object, encapsulating the conditional formats font, fill, borders, and other properties. Read-only.
         */
        getFormat(): ConditionalRangeFormat;

        /**
         * The criteria of the Top/Bottom conditional format.
         */
        getRule(): ConditionalTopBottomRule;
        setRule(rule: ConditionalTopBottomRule): void;
    }

    /**
     * Represents the the preset criteria conditional format such as above average, below average, unique values, contains blank, nonblank, error, and noerror.
     */
    interface PresetCriteriaConditionalFormat {
        /**
         * Returns a format object, encapsulating the conditional formats font, fill, borders, and other properties.
         */
        getFormat(): ConditionalRangeFormat;

        /**
         * The rule of the conditional format.
         */
        getRule(): ConditionalPresetCriteriaRule;
        setRule(rule: ConditionalPresetCriteriaRule): void;
    }

    /**
     * Represents a specific text conditional format.
     */
    interface TextConditionalFormat {
        /**
         * Returns a format object, encapsulating the conditional formats font, fill, borders, and other properties. Read-only.
         */
        getFormat(): ConditionalRangeFormat;

        /**
         * The rule of the conditional format.
         */
        getRule(): ConditionalTextComparisonRule;
        setRule(rule: ConditionalTextComparisonRule): void;
    }

    /**
     * Represents a cell value conditional format.
     */
    interface CellValueConditionalFormat {
        /**
         * Returns a format object, encapsulating the conditional formats font, fill, borders, and other properties.
         */
        getFormat(): ConditionalRangeFormat;

        /**
         * Represents the Rule object on this conditional format.
         */
        getRule(): ConditionalCellValueRule;
        setRule(rule: ConditionalCellValueRule): void;
    }

    /**
     * A format object encapsulating the conditional formats range's font, fill, borders, and other properties.
     */
    interface ConditionalRangeFormat {
        /**
         * Returns the fill object defined on the overall conditional format range. Read-only.
         */
        getFill(): ConditionalRangeFill;

        /**
         * Returns the font object defined on the overall conditional format range. Read-only.
         */
        getFont(): ConditionalRangeFont;

        getBorders(): ConditionalRangeBorder[];
        getConditionalRangeBorder(
            index: ConditionalRangeBorderIndex
        ): ConditionalRangeBorder;
    }

    /**
     * This object represents the font attributes (font style, color, etc.) for an object.
     */
    interface ConditionalRangeFont {
        /**
         * Represents the bold status of font.
         */
        getBold(): boolean;
        setBold(bold: boolean): void;

        /**
         * HTML color code representation of the text color. E.g. #FF0000 represents Red.
         */
        getColor(): string;
        setColor(color: string): void;

        /**
         * Represents the italic status of the font.
         */
        getItalic(): boolean;
        setItalic(italic: boolean): void;

        /**
         * Represents the strikethrough status of the font.
         */
        getStrikethrough(): boolean;
        setStrikethrough(strikethrough: boolean): void;

        /**
         * Type of underline applied to the font. See Excel.ConditionalRangeFontUnderlineStyle for details.
         */
        getUnderline(): ConditionalRangeFontUnderlineStyle;
        setUnderline(underline: ConditionalRangeFontUnderlineStyle): void;

        /**
         * Resets the font formats.
         */
        clear(): void;
    }

    /**
     * Represents the background of a conditional range object.
     */
    interface ConditionalRangeFill {
        /**
         * HTML color code representing the color of the fill, of the form #RRGGBB (e.g. "FFA500") or as a named HTML color (e.g. "orange").
         */
        getColor(): string;
        setColor(color: string): void;

        /**
         * Resets the fill.
         */
        clear(): void;
    }

    /**
     * Represents the border of an object.
     */
    interface ConditionalRangeBorder {
        /**
         * HTML color code representing the color of the border line, of the form #RRGGBB (e.g. "FFA500") or as a named HTML color (e.g. "orange").
         */
        getColor(): string;
        setColor(color: string): void;

        /**
         * Constant value that indicates the specific side of the border. See Excel.ConditionalRangeBorderIndex for details. Read-only.
         */
        getSideIndex(): ConditionalRangeBorderIndex;

        /**
         * One of the constants of line style specifying the line style for the border. See Excel.BorderLineStyle for details.
         */
        getStyle(): ConditionalRangeBorderLineStyle;
        setStyle(style: ConditionalRangeBorderLineStyle): void;
    }

    /**
     * An object encapsulating a style's format and other properties.
     */
    interface Style {
        /**
         * Indicates if text is automatically indented when the text alignment in a cell is set to equal distribution.
         */
        getAutoIndent(): boolean;
        setAutoIndent(autoIndent: boolean): void;

        /**
         * Indicates if the style is a built-in style.
         */
        getBuiltIn(): boolean;

        /**
         * The Fill of the style.
         */
        getFill(): RangeFill;

        /**
         * A Font object that represents the font of the style.
         */
        getFont(): RangeFont;

        /**
         * Indicates if the formula will be hidden when the worksheet is protected.
         */
        getFormulaHidden(): boolean;
        setFormulaHidden(formulaHidden: boolean): void;

        /**
         * Represents the horizontal alignment for the style. See Excel.HorizontalAlignment for details.
         */
        getHorizontalAlignment(): HorizontalAlignment;
        setHorizontalAlignment(horizontalAlignment: HorizontalAlignment): void;

        /**
         * Indicates if the style includes the AutoIndent, HorizontalAlignment, VerticalAlignment, WrapText, IndentLevel, and TextOrientation properties.
         */
        getIncludeAlignment(): boolean;
        setIncludeAlignment(includeAlignment: boolean): void;

        /**
         * Indicates if the style includes the Color, ColorIndex, LineStyle, and Weight border properties.
         */
        getIncludeBorder(): boolean;
        setIncludeBorder(includeBorder: boolean): void;

        /**
         * Indicates if the style includes the Background, Bold, Color, ColorIndex, FontStyle, Italic, Name, Size, Strikethrough, Subscript, Superscript, and Underline font properties.
         */
        getIncludeFont(): boolean;
        setIncludeFont(includeFont: boolean): void;

        /**
         * Indicates if the style includes the NumberFormat property.
         */
        getIncludeNumber(): boolean;
        setIncludeNumber(includeNumber: boolean): void;

        /**
         * Indicates if the style includes the Color, ColorIndex, InvertIfNegative, Pattern, PatternColor, and PatternColorIndex interior properties.
         */
        getIncludePatterns(): boolean;
        setIncludePatterns(includePatterns: boolean): void;

        /**
         * Indicates if the style includes the FormulaHidden and Locked protection properties.
         */
        getIncludeProtection(): boolean;
        setIncludeProtection(includeProtection: boolean): void;

        /**
         * An integer from 0 to 250 that indicates the indent level for the style.
         */
        getIndentLevel(): number;
        setIndentLevel(indentLevel: number): void;

        /**
         * Indicates if the object is locked when the worksheet is protected.
         */
        getLocked(): boolean;
        setLocked(locked: boolean): void;

        /**
         * The name of the style.
         */
        getName(): string;

        /**
         * The format code of the number format for the style.
         */
        getNumberFormat(): string;
        setNumberFormat(numberFormat: string): void;

        /**
         * The localized format code of the number format for the style.
         */
        getNumberFormatLocal(): string;
        setNumberFormatLocal(numberFormatLocal: string): void;

        /**
         * The reading order for the style.
         */
        getReadingOrder(): ReadingOrder;
        setReadingOrder(readingOrder: ReadingOrder): void;

        /**
         * Indicates if text automatically shrinks to fit in the available column width.
         */
        getShrinkToFit(): boolean;
        setShrinkToFit(shrinkToFit: boolean): void;

        /**
         * The text orientation for the style.
         */
        getTextOrientation(): number;
        setTextOrientation(textOrientation: number): void;

        /**
         * Represents the vertical alignment for the style. See Excel.VerticalAlignment for details.
         */
        getVerticalAlignment(): VerticalAlignment;
        setVerticalAlignment(verticalAlignment: VerticalAlignment): void;

        /**
         * Indicates if Microsoft Excel wraps the text in the object.
         */
        getWrapText(): boolean;
        setWrapText(wrapText: boolean): void;

        /**
         * Deletes this style.
         */
        delete(): void;

        getBorders(): RangeBorder[];
        getRangeBorder(index: BorderIndex): RangeBorder;
    }

    /**
     * Represents a TableStyle, which defines the style elements by region of the Table.
     */
    interface TableStyle {
        /**
         * Gets the name of the TableStyle.
         */
        getName(): string;
        setName(name: string): void;

        /**
         * Specifies whether this TableStyle object is read-only. Read-only.
         */
        getReadOnly(): boolean;

        /**
         * Deletes the TableStyle.
         */
        delete(): void;

        /**
         * Creates a duplicate of this TableStyle with copies of all the style elements.
         */
        duplicate(): TableStyle;
    }

    /**
     * Represents a PivotTable Style, which defines style elements by PivotTable region.
     */
    interface PivotTableStyle {
        /**
         * Gets the name of the PivotTableStyle.
         */
        getName(): string;
        setName(name: string): void;

        /**
         * Specifies whether this PivotTableStyle object is read-only. Read-only.
         */
        getReadOnly(): boolean;

        /**
         * Deletes the PivotTableStyle.
         */
        delete(): void;

        /**
         * Creates a duplicate of this PivotTableStyle with copies of all the style elements.
         */
        duplicate(): PivotTableStyle;
    }

    /**
     * Represents a Slicer Style, which defines style elements by region of the slicer.
     */
    interface SlicerStyle {
        /**
         * Gets the name of the SlicerStyle.
         */
        getName(): string;
        setName(name: string): void;

        /**
         * Specifies whether this SlicerStyle object is read-only. Read-only.
         */
        getReadOnly(): boolean;

        /**
         * Deletes the SlicerStyle.
         */
        delete(): void;

        /**
         * Creates a duplicate of this SlicerStyle with copies of all the style elements.
         */
        duplicate(): SlicerStyle;
    }

    /**
     * Represents a Timeline style, which defines style elements by region in the Timeline.
     */
    interface TimelineStyle {
        /**
         * Gets the name of the TimelineStyle.
         */
        getName(): string;
        setName(name: string): void;

        /**
         * Specifies whether this TimelineStyle object is read-only. Read-only.
         */
        getReadOnly(): boolean;

        /**
         * Deletes the TableStyle.
         */
        delete(): void;

        /**
         * Creates a duplicate of this TimelineStyle with copies of all the style elements.
         */
        duplicate(): TimelineStyle;
    }

    /**
     * Represents layout and print settings that are not dependent any printer-specific implementation. These settings include margins, orientation, page numbering, title rows, and print area.
     */
    interface PageLayout {
        /**
         * Gets or sets the worksheet's black and white print option.
         */
        getBlackAndWhite(): boolean;
        setBlackAndWhite(blackAndWhite: boolean): void;

        /**
         * Gets or sets the worksheet's bottom page margin to use for printing in points.
         */
        getBottomMargin(): number;
        setBottomMargin(bottomMargin: number): void;

        /**
         * Gets or sets the worksheet's center horizontally flag. This flag determines whether the worksheet will be centered horizontally when it's printed.
         */
        getCenterHorizontally(): boolean;
        setCenterHorizontally(centerHorizontally: boolean): void;

        /**
         * Gets or sets the worksheet's center vertically flag. This flag determines whether the worksheet will be centered vertically when it's printed.
         */
        getCenterVertically(): boolean;
        setCenterVertically(centerVertically: boolean): void;

        /**
         * Gets or sets the worksheet's draft mode option. If true the sheet will be printed without graphics.
         */
        getDraftMode(): boolean;
        setDraftMode(draftMode: boolean): void;

        /**
         * Gets or sets the worksheet's first page number to print. Null value represents "auto" page numbering.
         */
        getFirstPageNumber(): number | "";
        setFirstPageNumber(firstPageNumber: number | ""): void;

        /**
         * Gets or sets the worksheet's footer margin, in points, for use when printing.
         */
        getFooterMargin(): number;
        setFooterMargin(footerMargin: number): void;

        /**
         * Gets or sets the worksheet's header margin, in points, for use when printing.
         */
        getHeaderMargin(): number;
        setHeaderMargin(headerMargin: number): void;

        /**
         * Header and footer configuration for the worksheet.
         */
        getHeadersFooters(): HeaderFooterGroup;

        /**
         * Gets or sets the worksheet's left margin, in points, for use when printing.
         */
        getLeftMargin(): number;
        setLeftMargin(leftMargin: number): void;

        /**
         * Gets or sets the worksheet's orientation of the page.
         */
        getOrientation(): PageOrientation;
        setOrientation(orientation: PageOrientation): void;

        /**
         * Gets or sets the worksheet's paper size of the page.
         */
        getPaperSize(): PaperType;
        setPaperSize(paperSize: PaperType): void;

        /**
         * Gets or sets whether the worksheet's comments should be displayed when printing.
         */
        getPrintComments(): PrintComments;
        setPrintComments(printComments: PrintComments): void;

        /**
         * Gets or sets the worksheet's print errors option.
         */
        getPrintErrors(): PrintErrorType;
        setPrintErrors(printErrors: PrintErrorType): void;

        /**
         * Gets or sets the worksheet's print gridlines flag. This flag determines whether gridlines will be printed or not.
         */
        getPrintGridlines(): boolean;
        setPrintGridlines(printGridlines: boolean): void;

        /**
         * Gets or sets the worksheet's print headings flag. This flag determines whether headings will be printed or not.
         */
        getPrintHeadings(): boolean;
        setPrintHeadings(printHeadings: boolean): void;

        /**
         * Gets or sets the worksheet's page print order option. This specifies the order to use for processing the page number printed.
         */
        getPrintOrder(): PrintOrder;
        setPrintOrder(printOrder: PrintOrder): void;

        /**
         * Gets or sets the worksheet's right margin, in points, for use when printing.
         */
        getRightMargin(): number;
        setRightMargin(rightMargin: number): void;

        /**
         * Gets or sets the worksheet's top margin, in points, for use when printing.
         */
        getTopMargin(): number;
        setTopMargin(topMargin: number): void;

        /**
         * Gets or sets the worksheet's print zoom options.
         * The `PageLayoutZoomOptions` object must be set as a JSON object (use `x.zoom = {...}` instead of `x.zoom.scale = ...`).
         */
        getZoom(): PageLayoutZoomOptions;
        setZoom(zoom: PageLayoutZoomOptions): void;

        /**
         * Gets the RangeAreas object, comprising one or more rectangular ranges, that represents the print area for the worksheet. If there is no print area, a null object will be returned.
         */
        getPrintArea(): RangeAreas;

        /**
         * Gets the range object representing the title columns. If not set, this will return a null object.
         */
        getPrintTitleColumns(): Range;

        /**
         * Gets the range object representing the title rows. If not set, this will return a null object.
         */
        getPrintTitleRows(): Range;

        /**
         * Sets the worksheet's print area.
         */
        setPrintArea(printArea: Range | RangeAreas | string): void;

        /**
         * Sets the worksheet's page margins with units.
         */
        setPrintMargins(
            unit: PrintMarginUnit,
            marginOptions: PageLayoutMarginOptions
        ): void;

        /**
         * Sets the columns that contain the cells to be repeated at the left of each page of the worksheet for printing.
         */
        setPrintTitleColumns(printTitleColumns: Range | string): void;

        /**
         * Sets the rows that contain the cells to be repeated at the top of each page of the worksheet for printing.
         */
        setPrintTitleRows(printTitleRows: Range | string): void;
    }

    /**
     * no comment
     */
    interface HeaderFooter {
        /**
         * Gets or sets the center footer of the worksheet.
         * To apply font formatting or insert a variable value, use format codes specified here: https://msdn.microsoft.com/en-us/library/bb225426.aspx.
         */
        getCenterFooter(): string;
        setCenterFooter(centerFooter: string): void;

        /**
         * Gets or sets the center header of the worksheet.
         * To apply font formatting or insert a variable value, use format codes specified here: https://msdn.microsoft.com/en-us/library/bb225426.aspx.
         */
        getCenterHeader(): string;
        setCenterHeader(centerHeader: string): void;

        /**
         * Gets or sets the left footer of the worksheet.
         * To apply font formatting or insert a variable value, use format codes specified here: https://msdn.microsoft.com/en-us/library/bb225426.aspx.
         */
        getLeftFooter(): string;
        setLeftFooter(leftFooter: string): void;

        /**
         * Gets or sets the left header of the worksheet.
         * To apply font formatting or insert a variable value, use format codes specified here: https://msdn.microsoft.com/en-us/library/bb225426.aspx.
         */
        getLeftHeader(): string;
        setLeftHeader(leftHeader: string): void;

        /**
         * Gets or sets the right footer of the worksheet.
         * To apply font formatting or insert a variable value, use format codes specified here: https://msdn.microsoft.com/en-us/library/bb225426.aspx.
         */
        getRightFooter(): string;
        setRightFooter(rightFooter: string): void;

        /**
         * Gets or sets the right header of the worksheet.
         * To apply font formatting or insert a variable value, use format codes specified here: https://msdn.microsoft.com/en-us/library/bb225426.aspx.
         */
        getRightHeader(): string;
        setRightHeader(rightHeader: string): void;
    }

    /**
     * no comment
     */
    interface HeaderFooterGroup {
        /**
         * The general header/footer, used for all pages unless even/odd or first page is specified.
         */
        getDefaultForAllPages(): HeaderFooter;

        /**
         * The header/footer to use for even pages, odd header/footer needs to be specified for odd pages.
         */
        getEvenPages(): HeaderFooter;

        /**
         * The first page header/footer, for all other pages general or even/odd is used.
         */
        getFirstPage(): HeaderFooter;

        /**
         * The header/footer to use for odd pages, even header/footer needs to be specified for even pages.
         */
        getOddPages(): HeaderFooter;

        /**
         * Gets or sets the state of which headers/footers are set. See Excel.HeaderFooterState for details.
         */
        getState(): HeaderFooterState;
        setState(state: HeaderFooterState): void;

        /**
         * Gets or sets a flag indicating if headers/footers are aligned with the page margins set in the page layout options for the worksheet.
         */
        getUseSheetMargins(): boolean;
        setUseSheetMargins(useSheetMargins: boolean): void;

        /**
         * Gets or sets a flag indicating if headers/footers should be scaled by the page percentage scale set in the page layout options for the worksheet.
         */
        getUseSheetScale(): boolean;
        setUseSheetScale(useSheetScale: boolean): void;
    }

    /**
     * no comment
     */
    interface PageBreak {
        /**
         * Represents the column index for the page break
         */
        getColumnIndex(): number;

        /**
         * Deletes a page break object.
         */
        delete(): void;

        /**
         * Gets the first cell after the page break.
         */
        getCellAfterBreak(): Range;
    }

    /**
     * Represents a comment in the workbook.
     */
    interface Comment {
        /**
         * Gets the email of the comment's author.
         */
        getAuthorEmail(): string;

        /**
         * Gets the name of the comment's author.
         */
        getAuthorName(): string;

        /**
         * Gets or sets the comment's content. The string is plain text.
         */
        getContent(): string;
        setContent(content: string): void;

        /**
         * Gets the creation time of the comment. Returns null if the comment was converted from a note, since the comment does not have a creation date.
         */
        getCreationDate(): Date;

        /**
         * Represents the comment identifier. Read-only.
         */
        getId(): string;

        /**
         * Gets the entities (e.g. people) that are mentioned in comments.
         */
        getMentions(): CommentMention[];

        /**
         * Gets the rich comment content (e.g. mentions in comments). This string is not meant to be displayed to end-users. Your add-in should only use this to parse rich comment content.
         */
        getRichContent(): string;

        /**
         * Deletes the comment and all the connected replies.
         */
        delete(): void;

        /**
         * Gets the cell where this comment is located.
         */
        getLocation(): Range;

        /**
         * Updates the comment content with a specially formatted string and a list of mentions.
         */
        updateMentions(contentWithMentions: CommentRichContent): void;

        getReplies(): CommentReply[];
        getCommentReply(commentReplyId: string): CommentReply;
    }

    /**
     * Represents a comment reply in the workbook.
     */
    interface CommentReply {
        /**
         * Gets the email of the comment reply's author.
         */
        getAuthorEmail(): string;

        /**
         * Gets the name of the comment reply's author.
         */
        getAuthorName(): string;

        /**
         * Gets or sets the comment reply's content. The string is plain text.
         */
        getContent(): string;
        setContent(content: string): void;

        /**
         * Gets the creation time of the comment reply.
         */
        getCreationDate(): Date;

        /**
         * Represents the comment reply identifier. Read-only.
         */
        getId(): string;

        /**
         * Gets the entities (e.g. people) that are mentioned in comments.
         */
        getMentions(): CommentMention[];

        /**
         * Gets the rich comment content (e.g. mentions in comments). This string is not meant to be displayed to end-users. Your add-in should only use this to parse rich comment content.
         */
        getRichContent(): string;

        /**
         * Deletes the comment reply.
         */
        delete(): void;

        /**
         * Gets the cell where this comment reply is located.
         */
        getLocation(): Range;

        /**
         * Gets the parent comment of this reply.
         */
        getParentComment(): Comment;

        /**
         * Updates the comment content with a specially formatted string and a list of mentions.
         */
        updateMentions(contentWithMentions: CommentRichContent): void;
    }

    /**
     * Represents a generic shape object in the worksheet. A shape could be a geometric shape, a line, a group of shapes, etc.
     * To learn more about the shape object model, read {@link https://docs.microsoft.com/office/dev/add-ins/excel/excel-add-ins-shapes | Work with shapes using the Excel JavaScript API}.
     */
    interface Shape {
        /**
         * Returns or sets the alternative description text for a Shape object.
         */
        getAltTextDescription(): string;
        setAltTextDescription(altTextDescription: string): void;

        /**
         * Returns or sets the alternative title text for a Shape object.
         */
        getAltTextTitle(): string;
        setAltTextTitle(altTextTitle: string): void;

        /**
         * Returns the number of connection sites on this shape. Read-only.
         */
        getConnectionSiteCount(): number;

        /**
         * Returns the fill formatting of this shape. Read-only.
         */
        getFill(): ShapeFill;

        /**
         * Returns the geometric shape associated with the shape. An error will be thrown if the shape type is not "GeometricShape".
         */
        getGeometricShape(): GeometricShape;

        /**
         * Represents the geometric shape type of this geometric shape. See Excel.GeometricShapeType for details. Returns null if the shape type is not "GeometricShape".
         */
        getGeometricShapeType(): GeometricShapeType;
        setGeometricShapeType(geometricShapeType: GeometricShapeType): void;

        /**
         * Returns the shape group associated with the shape. An error will be thrown if the shape type is not "GroupShape".
         */
        getGroup(): ShapeGroup;

        /**
         * Represents the height, in points, of the shape.
         * Throws an invalid argument exception when set with a negative value or zero as input.
         */
        getHeight(): number;
        setHeight(height: number): void;

        /**
         * Represents the shape identifier. Read-only.
         */
        getId(): string;

        /**
         * Returns the image associated with the shape. An error will be thrown if the shape type is not "Image".
         */
        getImage(): Image;

        /**
         * The distance, in points, from the left side of the shape to the left side of the worksheet.
         * Throws an invalid argument exception when set with a negative value as input.
         */
        getLeft(): number;
        setLeft(left: number): void;

        /**
         * Represents the level of the specified shape. For example, a level of 0 means that the shape is not part of any groups, a level of 1 means the shape is part of a top-level group, and a level of 2 means the shape is part of a sub-group of the top level.
         */
        getLevel(): number;

        /**
         * Returns the line associated with the shape. An error will be thrown if the shape type is not "Line".
         */
        getLine(): Line;

        /**
         * Returns the line formatting of this shape. Read-only.
         */
        getLineFormat(): ShapeLineFormat;

        /**
         * Specifies whether or not the aspect ratio of this shape is locked.
         */
        getLockAspectRatio(): boolean;
        setLockAspectRatio(lockAspectRatio: boolean): void;

        /**
         * Represents the name of the shape.
         */
        getName(): string;
        setName(name: string): void;

        /**
         * Represents how the object is attached to the cells below it.
         */
        getPlacement(): Placement;
        setPlacement(placement: Placement): void;

        /**
         * Represents the rotation, in degrees, of the shape.
         */
        getRotation(): number;
        setRotation(rotation: number): void;

        /**
         * Returns the text frame object of this shape. Read only.
         */
        getTextFrame(): TextFrame;

        /**
         * The distance, in points, from the top edge of the shape to the top edge of the worksheet.
         * Throws an invalid argument exception when set with a negative value as input.
         */
        getTop(): number;
        setTop(top: number): void;

        /**
         * Returns the type of this shape. See Excel.ShapeType for details. Read-only.
         */
        getType(): ShapeType;

        /**
         * Represents the visibility of this shape.
         */
        getVisible(): boolean;
        setVisible(visible: boolean): void;

        /**
         * Represents the width, in points, of the shape.
         * Throws an invalid argument exception when set with a negative value or zero as input.
         */
        getWidth(): number;
        setWidth(width: number): void;

        /**
         * Returns the position of the specified shape in the z-order, with 0 representing the bottom of the order stack. Read-only.
         */
        getZOrderPosition(): number;

        /**
         * Copies and pastes a Shape object.
         * The pasted shape is copied to the same pixel location as this shape.
         */
        copyTo(destinationSheet?: Worksheet | string): Shape;

        /**
         * Removes the shape from the worksheet.
         */
        delete(): void;

        /**
         * Converts the shape to an image and returns the image as a base64-encoded string. The DPI is 96. The only supported formats are `Excel.PictureFormat.BMP`, `Excel.PictureFormat.PNG`, `Excel.PictureFormat.JPEG`, and `Excel.PictureFormat.GIF`.
         */
        getAsImage(format: PictureFormat): string;

        /**
         * Moves the shape horizontally by the specified number of points.
         */
        incrementLeft(increment: number): void;

        /**
         * Rotates the shape clockwise around the z-axis by the specified number of degrees.
         * Use the `rotation` property to set the absolute rotation of the shape.
         */
        incrementRotation(increment: number): void;

        /**
         * Moves the shape vertically by the specified number of points.
         */
        incrementTop(increment: number): void;

        /**
         * Scales the height of the shape by a specified factor. For images, you can indicate whether you want to scale the shape relative to the original or the current size. Shapes other than pictures are always scaled relative to their current height.
         */
        scaleHeight(
            scaleFactor: number,
            scaleType: ShapeScaleType,
            scaleFrom?: ShapeScaleFrom
        ): void;

        /**
         * Scales the width of the shape by a specified factor. For images, you can indicate whether you want to scale the shape relative to the original or the current size. Shapes other than pictures are always scaled relative to their current width.
         */
        scaleWidth(
            scaleFactor: number,
            scaleType: ShapeScaleType,
            scaleFrom?: ShapeScaleFrom
        ): void;

        /**
         * Moves the specified shape up or down the collection's z-order, which shifts it in front of or behind other shapes.
         */
        setZOrder(position: ShapeZOrder): void;
    }

    /**
     * Represents a geometric shape inside a worksheet. A geometric shape can be a rectangle, block arrow, equation symbol, flowchart item, star, banner, callout, or any other basic shape in Excel.
     */
    interface GeometricShape {
        /**
         * Returns the shape identifier. Read-only.
         */
        getId(): string;
    }

    /**
     * Represents an image in the worksheet. To get the corresponding Shape object, use Image.shape.
     */
    interface Image {
        /**
         * Represents the shape identifier for the image object. Read-only.
         */
        getId(): string;

        /**
         * Returns the format of the image. Read-only.
         */
        getFormat(): PictureFormat;
    }

    /**
     * Represents a shape group inside a worksheet. To get the corresponding Shape object, use `ShapeGroup.shape`.
     */
    interface ShapeGroup {
        /**
         * Represents the shape identifier. Read-only.
         */
        getId(): string;

        /**
         * Ungroups any grouped shapes in the specified shape group.
         */
        ungroup(): void;
    }

    /**
     * Represents a line inside a worksheet. To get the corresponding Shape object, use `Line.shape`.
     */
    interface Line {
        /**
         * Represents the length of the arrowhead at the beginning of the specified line.
         */
        getBeginArrowheadLength(): ArrowheadLength;
        setBeginArrowheadLength(beginArrowheadLength: ArrowheadLength): void;

        /**
         * Represents the style of the arrowhead at the beginning of the specified line.
         */
        getBeginArrowheadStyle(): ArrowheadStyle;
        setBeginArrowheadStyle(beginArrowheadStyle: ArrowheadStyle): void;

        /**
         * Represents the width of the arrowhead at the beginning of the specified line.
         */
        getBeginArrowheadWidth(): ArrowheadWidth;
        setBeginArrowheadWidth(beginArrowheadWidth: ArrowheadWidth): void;

        /**
         * Represents the connection site to which the beginning of a connector is connected. Read-only. Returns null when the beginning of the line is not attached to any shape.
         */
        getBeginConnectedSite(): number;

        /**
         * Represents the length of the arrowhead at the end of the specified line.
         */
        getEndArrowheadLength(): ArrowheadLength;
        setEndArrowheadLength(endArrowheadLength: ArrowheadLength): void;

        /**
         * Represents the style of the arrowhead at the end of the specified line.
         */
        getEndArrowheadStyle(): ArrowheadStyle;
        setEndArrowheadStyle(endArrowheadStyle: ArrowheadStyle): void;

        /**
         * Represents the width of the arrowhead at the end of the specified line.
         */
        getEndArrowheadWidth(): ArrowheadWidth;
        setEndArrowheadWidth(endArrowheadWidth: ArrowheadWidth): void;

        /**
         * Represents the connection site to which the end of a connector is connected. Read-only. Returns null when the end of the line is not attached to any shape.
         */
        getEndConnectedSite(): number;

        /**
         * Represents the shape identifier. Read-only.
         */
        getId(): string;

        /**
         * Specifies whether or not the beginning of the specified line is connected to a shape. Read-only.
         */
        getIsBeginConnected(): boolean;

        /**
         * Specifies whether or not the end of the specified line is connected to a shape. Read-only.
         */
        getIsEndConnected(): boolean;

        /**
         * Represents the connector type for the line.
         */
        getConnectorType(): ConnectorType;
        setConnectorType(connectorType: ConnectorType): void;

        /**
         * Attaches the beginning of the specified connector to a specified shape.
         */
        connectBeginShape(shape: Shape, connectionSite: number): void;

        /**
         * Attaches the end of the specified connector to a specified shape.
         */
        connectEndShape(shape: Shape, connectionSite: number): void;

        /**
         * Detaches the beginning of the specified connector from a shape.
         */
        disconnectBeginShape(): void;

        /**
         * Detaches the end of the specified connector from a shape.
         */
        disconnectEndShape(): void;
    }

    /**
     * Represents the fill formatting of a shape object.
     */
    interface ShapeFill {
        /**
         * Represents the shape fill foreground color in HTML color format, of the form #RRGGBB (e.g. "FFA500") or as a named HTML color (e.g. "orange")
         */
        getForegroundColor(): string;
        setForegroundColor(foregroundColor: string): void;

        /**
         * Returns or sets the transparency percentage of the fill as a value from 0.0 (opaque) through 1.0 (clear). Returns null if the shape type does not support transparency or the shape fill has inconsistent transparency, such as with a gradient fill type.
         */
        getTransparency(): number;
        setTransparency(transparency: number): void;

        /**
         * Returns the fill type of the shape. Read-only. See Excel.ShapeFillType for details.
         */
        getType(): ShapeFillType;

        /**
         * Clears the fill formatting of this shape.
         */
        clear(): void;

        /**
         * Sets the fill formatting of the shape to a uniform color. This changes the fill type to "Solid".
         */
        setSolidColor(color: string): void;
    }

    /**
     * Represents the line formatting for the shape object. For images and geometric shapes, line formatting represents the border of the shape.
     */
    interface ShapeLineFormat {
        /**
         * Represents the line color in HTML color format, of the form #RRGGBB (e.g. "FFA500") or as a named HTML color (e.g. "orange").
         */
        getColor(): string;
        setColor(color: string): void;

        /**
         * Represents the line style of the shape. Returns null when the line is not visible or there are inconsistent dash styles. See Excel.ShapeLineStyle for details.
         */
        getDashStyle(): ShapeLineDashStyle;
        setDashStyle(dashStyle: ShapeLineDashStyle): void;

        /**
         * Represents the line style of the shape. Returns null when the line is not visible or there are inconsistent styles. See Excel.ShapeLineStyle for details.
         */
        getStyle(): ShapeLineStyle;
        setStyle(style: ShapeLineStyle): void;

        /**
         * Represents the degree of transparency of the specified line as a value from 0.0 (opaque) through 1.0 (clear). Returns null when the shape has inconsistent transparencies.
         */
        getTransparency(): number;
        setTransparency(transparency: number): void;

        /**
         * Represents whether or not the line formatting of a shape element is visible. Returns null when the shape has inconsistent visibilities.
         */
        getVisible(): boolean;
        setVisible(visible: boolean): void;

        /**
         * Represents the weight of the line, in points. Returns null when the line is not visible or there are inconsistent line weights.
         */
        getWeight(): number;
        setWeight(weight: number): void;
    }

    /**
     * Represents the text frame of a shape object.
     */
    interface TextFrame {
        /**
         * Gets or sets the automatic sizing settings for the text frame. A text frame can be set to automatically fit the text to the text frame, to automatically fit the text frame to the text, or not perform any automatic sizing.
         */
        getAutoSizeSetting(): ShapeAutoSize;
        setAutoSizeSetting(autoSizeSetting: ShapeAutoSize): void;

        /**
         * Represents the bottom margin, in points, of the text frame.
         */
        getBottomMargin(): number;
        setBottomMargin(bottomMargin: number): void;

        /**
         * Specifies whether the text frame contains text.
         */
        getHasText(): boolean;

        /**
         * Represents the horizontal alignment of the text frame. See Excel.ShapeTextHorizontalAlignment for details.
         */
        getHorizontalAlignment(): ShapeTextHorizontalAlignment;
        setHorizontalAlignment(
            horizontalAlignment: ShapeTextHorizontalAlignment
        ): void;

        /**
         * Represents the horizontal overflow behavior of the text frame. See Excel.ShapeTextHorizontalOverflow for details.
         */
        getHorizontalOverflow(): ShapeTextHorizontalOverflow;
        setHorizontalOverflow(
            horizontalOverflow: ShapeTextHorizontalOverflow
        ): void;

        /**
         * Represents the left margin, in points, of the text frame.
         */
        getLeftMargin(): number;
        setLeftMargin(leftMargin: number): void;

        /**
         * Represents the angle to which the text is oriented for the text frame. See Excel.ShapeTextOrientation for details.
         */
        getOrientation(): ShapeTextOrientation;
        setOrientation(orientation: ShapeTextOrientation): void;

        /**
         * Represents the reading order of the text frame, either left-to-right or right-to-left. See Excel.ShapeTextReadingOrder for details.
         */
        getReadingOrder(): ShapeTextReadingOrder;
        setReadingOrder(readingOrder: ShapeTextReadingOrder): void;

        /**
         * Represents the right margin, in points, of the text frame.
         */
        getRightMargin(): number;
        setRightMargin(rightMargin: number): void;

        /**
         * Represents the text that is attached to a shape in the text frame, and properties and methods for manipulating the text. See Excel.TextRange for details.
         */
        getTextRange(): TextRange;

        /**
         * Represents the top margin, in points, of the text frame.
         */
        getTopMargin(): number;
        setTopMargin(topMargin: number): void;

        /**
         * Represents the vertical alignment of the text frame. See Excel.ShapeTextVerticalAlignment for details.
         */
        getVerticalAlignment(): ShapeTextVerticalAlignment;
        setVerticalAlignment(
            verticalAlignment: ShapeTextVerticalAlignment
        ): void;

        /**
         * Represents the vertical overflow behavior of the text frame. See Excel.ShapeTextVerticalOverflow for details.
         */
        getVerticalOverflow(): ShapeTextVerticalOverflow;
        setVerticalOverflow(verticalOverflow: ShapeTextVerticalOverflow): void;

        /**
         * Deletes all the text in the text frame.
         */
        deleteText(): void;
    }

    /**
     * Contains the text that is attached to a shape, in addition to properties and methods for manipulating the text.
     */
    interface TextRange {
        /**
         * Returns a ShapeFont object that represents the font attributes for the text range. Read-only.
         */
        getFont(): ShapeFont;

        /**
         * Represents the plain text content of the text range.
         */
        getText(): string;
        setText(text: string): void;

        /**
         * Returns a TextRange object for the substring in the given range.
         */
        getSubstring(start: number, length?: number): TextRange;
    }

    /**
     * Represents the font attributes, such as font name, font size, and color, for a shape's TextRange object.
     */
    interface ShapeFont {
        /**
         * Represents the bold status of font. Returns null the TextRange includes both bold and non-bold text fragments.
         */
        getBold(): boolean;
        setBold(bold: boolean): void;

        /**
         * The HTML color code representation of the text color (e.g. "#FF0000" represents red). Returns null if the TextRange includes text fragments with different colors.
         */
        getColor(): string;
        setColor(color: string): void;

        /**
         * Represents the italic status of font. Returns null if the TextRange includes both italic and non-italic text fragments.
         */
        getItalic(): boolean;
        setItalic(italic: boolean): void;

        /**
         * Represents font name (e.g. "Calibri"). If the text is Complex Script or East Asian language, this is the corresponding font name; otherwise it is the Latin font name.
         */
        getName(): string;
        setName(name: string): void;

        /**
         * Represents font size in points (e.g. 11). Returns null if the TextRange includes text fragments with different font sizes.
         */
        getSize(): number;
        setSize(size: number): void;

        /**
         * Type of underline applied to the font. Returns null if the TextRange includes text fragments with different underline styles. See Excel.ShapeFontUnderlineStyle for details.
         */
        getUnderline(): ShapeFontUnderlineStyle;
        setUnderline(underline: ShapeFontUnderlineStyle): void;
    }

    /**
     * Represents a slicer object in the workbook.
     */
    interface Slicer {
        /**
         * Represents the caption of slicer.
         */
        getCaption(): string;
        setCaption(caption: string): void;

        /**
         * Represents the height, in points, of the slicer.
         * Throws an "The argument is invalid or missing or has an incorrect format." exception when set with negative value or zero as input.
         */
        getHeight(): number;
        setHeight(height: number): void;

        /**
         * Represents the unique id of slicer. Read-only.
         */
        getId(): string;

        /**
         * True if all filters currently applied on the slicer are cleared.
         */
        getIsFilterCleared(): boolean;

        /**
         * Represents the distance, in points, from the left side of the slicer to the left of the worksheet.
         * Throws an "The argument is invalid or missing or has an incorrect format." exception when set with negative value as input.
         */
        getLeft(): number;
        setLeft(left: number): void;

        /**
         * Represents the name of slicer.
         */
        getName(): string;
        setName(name: string): void;

        /**
         * Represents the sort order of the items in the slicer. Possible values are: "DataSourceOrder", "Ascending", "Descending".
         */
        getSortBy(): SlicerSortType;
        setSortBy(sortBy: SlicerSortType): void;

        /**
         * Constant value that represents the Slicer style. Possible values are: "SlicerStyleLight1" through "SlicerStyleLight6", "TableStyleOther1" through "TableStyleOther2", "SlicerStyleDark1" through "SlicerStyleDark6". A custom user-defined style present in the workbook can also be specified.
         */
        getStyle(): string;
        setStyle(style: string): void;

        /**
         * Represents the distance, in points, from the top edge of the slicer to the top of the worksheet.
         * Throws an "The argument is invalid or missing or has an incorrect format." exception when set with negative value as input.
         */
        getTop(): number;
        setTop(top: number): void;

        /**
         * Represents the width, in points, of the slicer.
         * Throws an "The argument is invalid or missing or has an incorrect format." exception when set with negative value or zero as input.
         */
        getWidth(): number;
        setWidth(width: number): void;

        /**
         * Represents the worksheet containing the slicer. Read-only.
         */
        getWorksheet(): Worksheet;

        /**
         * Clears all the filters currently applied on the slicer.
         */
        clearFilters(): void;

        /**
         * Deletes the slicer.
         */
        delete(): void;

        /**
         * Returns an array of selected items' keys. Read-only.
         */
        getSelectedItems(): string[];

        /**
         * Selects slicer items based on their keys. The previous selections are cleared.
         * All items will be selected by default if the array is empty.
         */
        selectItems(items?: string[]): void;

        getSlicerItems(): SlicerItem[];
        getSlicerItem(key: string): SlicerItem | undefined;
    }

    /**
     * Represents a slicer item in a slicer.
     */
    interface SlicerItem {
        /**
         * True if the slicer item has data.
         */
        getHasData(): boolean;

        /**
         * True if the slicer item is selected.
         * Setting this value will not clear other SlicerItems' selected state.
         * By default, if the slicer item is the only one selected, when it is deselected, all items will be selected.
         */
        getIsSelected(): boolean;
        setIsSelected(isSelected: boolean): void;

        /**
         * Represents the unique value representing the slicer item.
         */
        getKey(): string;

        /**
         * Represents the title displayed in the UI.
         */
        getName(): string;
    }

    //
    // Interface
    //

    interface WorksheetProtectionOptions {
        allowAutoFilter?: boolean;
        allowDeleteColumns?: boolean;
        allowDeleteRows?: boolean;
        allowEditObjects?: boolean;
        allowEditScenarios?: boolean;
        allowFormatCells?: boolean;
        allowFormatColumns?: boolean;
        allowFormatRows?: boolean;
        allowInsertColumns?: boolean;
        allowInsertHyperlinks?: boolean;
        allowInsertRows?: boolean;
        allowPivotTables?: boolean;
        allowSort?: boolean;
        selectionMode?: ProtectionSelectionMode;
    }

    interface RangeReference {
        address: string;
    }

    interface RangeHyperlink {
        address?: string;
        documentReference?: string;
        screenTip?: string;
        textToDisplay?: string;
    }

    interface SearchCriteria {
        completeMatch?: boolean;
        matchCase?: boolean;
        searchDirection?: SearchDirection;
    }

    interface WorksheetSearchCriteria {
        completeMatch?: boolean;
        matchCase?: boolean;
    }

    interface ReplaceCriteria {
        completeMatch?: boolean;
        matchCase?: boolean;
    }

    interface CellPropertiesFillLoadOptions {
        color?: boolean;
        pattern?: boolean;
        patternColor?: boolean;
        patternTintAndShade?: boolean;
        tintAndShade?: boolean;
    }

    interface CellPropertiesFontLoadOptions {
        bold?: boolean;
        color?: boolean;
        italic?: boolean;
        name?: boolean;
        size?: boolean;
        strikethrough?: boolean;
        subscript?: boolean;
        superscript?: boolean;
        tintAndShade?: boolean;
        underline?: boolean;
    }

    interface CellPropertiesBorderLoadOptions {
        color?: boolean;
        style?: boolean;
        tintAndShade?: boolean;
        weight?: boolean;
    }

    interface CellPropertiesProtection {
        formulaHidden?: boolean;
        locked?: boolean;
    }

    interface CellPropertiesFill {
        color?: string;
        pattern?: FillPattern;
        patternColor?: string;
        patternTintAndShade?: number;
        tintAndShade?: number;
    }

    interface CellPropertiesFont {
        bold?: boolean;
        color?: string;
        italic?: boolean;
        name?: string;
        size?: number;
        strikethrough?: boolean;
        subscript?: boolean;
        superscript?: boolean;
        tintAndShade?: number;
        underline?: RangeUnderlineStyle;
    }

    interface CellBorderCollection {
        bottom?: CellBorder;
        diagonalDown?: CellBorder;
        diagonalUp?: CellBorder;
        horizontal?: CellBorder;
        left?: CellBorder;
        right?: CellBorder;
        top?: CellBorder;
        vertical?: CellBorder;
    }

    interface CellBorder {
        color?: string;
        style?: BorderLineStyle;
        tintAndShade?: number;
        weight?: BorderWeight;
    }

    interface DataValidationRule {
        custom?: CustomDataValidation;
        date?: DateTimeDataValidation;
        decimal?: BasicDataValidation;
        list?: ListDataValidation;
        textLength?: BasicDataValidation;
        time?: DateTimeDataValidation;
        wholeNumber?: BasicDataValidation;
    }

    interface BasicDataValidation {
        formula1: string | number | Range;
        formula2?: string | number | Range;
        operator: DataValidationOperator;
    }

    interface DateTimeDataValidation {
        formula1: string | Date | Range;
        formula2?: string | Date | Range;
        operator: DataValidationOperator;
    }

    interface ListDataValidation {
        inCellDropDown: boolean;
        source: string | Range;
    }

    interface CustomDataValidation {
        formula: string;
    }

    interface DataValidationErrorAlert {
        message: string;
        showAlert: boolean;
        style: DataValidationAlertStyle;
        title: string;
    }

    interface SortField {
        ascending?: boolean;
        color?: string;
        dataOption?: SortDataOption;
        icon?: Icon;
        key: number;
        sortOn?: SortOn;
        subField?: string;
    }

    interface FilterCriteria {
        color?: string;
        criterion1?: string;
        criterion2?: string;
        dynamicCriteria?: DynamicFilterCriteria;
        filterOn: FilterOn;
        icon?: Icon;
        operator?: FilterOperator;
        subField?: string;
        values?: Array<string | FilterDatetime>;
    }

    interface FilterDatetime {
        date: string;
        specificity: FilterDatetimeSpecificity;
    }

    interface Icon {
        index: number;
        set: IconSet;
    }

    interface ShowAsRule {
        baseField?: PivotField;
        baseItem?: PivotItem;
        calculation: ShowAsCalculation;
    }

    interface Subtotals {
        automatic?: boolean;
        average?: boolean;
        count?: boolean;
        countNumbers?: boolean;
        max?: boolean;
        min?: boolean;
        product?: boolean;
        standardDeviation?: boolean;
        standardDeviationP?: boolean;
        sum?: boolean;
        variance?: boolean;
        varianceP?: boolean;
    }

    interface ConditionalDataBarRule {
        formula?: string;
        type: ConditionalFormatRuleType;
    }

    interface ConditionalIconCriterion {
        customIcon?: Icon;
        formula: string;
        operator: ConditionalIconCriterionOperator;
        type: ConditionalFormatIconRuleType;
    }

    interface ConditionalColorScaleCriteria {
        maximum: ConditionalColorScaleCriterion;
        midpoint?: ConditionalColorScaleCriterion;
        minimum: ConditionalColorScaleCriterion;
    }

    interface ConditionalColorScaleCriterion {
        color?: string;
        formula?: string;
        type: ConditionalFormatColorCriterionType;
    }

    interface ConditionalTopBottomRule {
        rank: number;
        type: ConditionalTopBottomCriterionType;
    }

    interface ConditionalPresetCriteriaRule {
        criterion: ConditionalFormatPresetCriterion;
    }

    interface ConditionalTextComparisonRule {
        operator: ConditionalTextOperator;
        text: string;
    }

    interface ConditionalCellValueRule {
        formula1: string;
        formula2?: string;
        operator: ConditionalCellValueOperator;
    }

    interface PageLayoutZoomOptions {
        horizontalFitToPages?: number;
        scale?: number;
        verticalFitToPages?: number;
    }

    interface PageLayoutMarginOptions {
        bottom?: number;
        footer?: number;
        header?: number;
        left?: number;
        right?: number;
        top?: number;
    }

    interface CommentMention {
        email: string;
        id: number;
        name: string;
    }

    interface CommentRichContent {
        mentions?: CommentMention[];
        richContent: string;
    }

    //
    // Enum
    //

    enum PivotFilterTopBottomCriterion {
        invalid,
        topItems,
        topPercent,
        topSum,
        bottomItems,
        bottomPercent,
        bottomSum,
    }

    enum SortBy {
        ascending,
        descending,
    }

    enum AggregationFunction {
        unknown,
        automatic,
        sum,
        count,
        average,
        max,
        min,
        product,
        countNumbers,
        standardDeviation,
        standardDeviationP,
        variance,
        varianceP,
    }

    enum ShowAsCalculation {
        unknown,
        none,
        percentOfGrandTotal,
        percentOfRowTotal,
        percentOfColumnTotal,
        percentOfParentRowTotal,
        percentOfParentColumnTotal,
        percentOfParentTotal,
        percentOf,
        runningTotal,
        percentRunningTotal,
        differenceFrom,
        percentDifferenceFrom,
        rankAscending,
        rankDecending,
        index,
    }

    enum PivotAxis {
        unknown,
        row,
        column,
        data,
        filter,
    }

    enum ChartAxisType {
        invalid,
        category,
        value,
        series,
    }

    enum ChartAxisGroup {
        primary,
        secondary,
    }

    enum ChartAxisScaleType {
        linear,
        logarithmic,
    }

    enum ChartAxisPosition {
        automatic,
        maximum,
        minimum,
        custom,
    }

    enum ChartAxisTickMark {
        none,
        cross,
        inside,
        outside,
    }

    enum CalculationState {
        done,
        calculating,
        pending,
    }

    enum ChartAxisTickLabelPosition {
        nextToAxis,
        high,
        low,
        none,
    }

    enum ChartAxisDisplayUnit {
        none,
        hundreds,
        thousands,
        tenThousands,
        hundredThousands,
        millions,
        tenMillions,
        hundredMillions,
        billions,
        trillions,
        custom,
    }

    enum ChartAxisTimeUnit {
        days,
        months,
        years,
    }

    enum ChartBoxQuartileCalculation {
        inclusive,
        exclusive,
    }

    enum ChartAxisCategoryType {
        automatic,
        textAxis,
        dateAxis,
    }

    enum ChartBinType {
        category,
        auto,
        binWidth,
        binCount,
    }

    enum ChartLineStyle {
        none,
        continuous,
        dash,
        dashDot,
        dashDotDot,
        dot,
        grey25,
        grey50,
        grey75,
        automatic,
        roundDot,
    }

    enum ChartDataLabelPosition {
        invalid,
        none,
        center,
        insideEnd,
        insideBase,
        outsideEnd,
        left,
        right,
        top,
        bottom,
        bestFit,
        callout,
    }

    enum ChartErrorBarsInclude {
        both,
        minusValues,
        plusValues,
    }

    enum ChartErrorBarsType {
        fixedValue,
        percent,
        stDev,
        stError,
        custom,
    }

    enum ChartMapAreaLevel {
        automatic,
        dataOnly,
        city,
        county,
        state,
        country,
        continent,
        world,
    }

    enum ChartGradientStyle {
        twoPhaseColor,
        threePhaseColor,
    }

    enum ChartGradientStyleType {
        extremeValue,
        number,
        percent,
    }

    enum ChartTitlePosition {
        automatic,
        top,
        bottom,
        left,
        right,
    }

    enum ChartLegendPosition {
        invalid,
        top,
        bottom,
        left,
        right,
        corner,
        custom,
    }

    enum ChartMarkerStyle {
        invalid,
        automatic,
        none,
        square,
        diamond,
        triangle,
        x,
        star,
        dot,
        dash,
        circle,
        plus,
        picture,
    }

    enum ChartPlotAreaPosition {
        automatic,
        custom,
    }

    enum ChartMapLabelStrategy {
        none,
        bestFit,
        showAll,
    }

    enum ChartMapProjectionType {
        automatic,
        mercator,
        miller,
        robinson,
        albers,
    }

    enum ChartParentLabelStrategy {
        none,
        banner,
        overlapping,
    }

    enum ChartSeriesBy {
        auto,
        columns,
        rows,
    }

    enum ChartTextHorizontalAlignment {
        center,
        left,
        right,
        justify,
        distributed,
    }

    enum ChartTextVerticalAlignment {
        center,
        bottom,
        top,
        justify,
        distributed,
    }

    enum ChartTickLabelAlignment {
        center,
        left,
        right,
    }

    enum ChartType {
        invalid,
        columnClustered,
        columnStacked,
        columnStacked100,
        barClustered,
        barStacked,
        barStacked100,
        lineStacked,
        lineStacked100,
        lineMarkers,
        lineMarkersStacked,
        lineMarkersStacked100,
        pieOfPie,
        pieExploded,
        barOfPie,
        xyscatterSmooth,
        xyscatterSmoothNoMarkers,
        xyscatterLines,
        xyscatterLinesNoMarkers,
        areaStacked,
        areaStacked100,
        doughnutExploded,
        radarMarkers,
        radarFilled,
        surface,
        surfaceWireframe,
        surfaceTopView,
        surfaceTopViewWireframe,
        bubble,
        bubble3DEffect,
        stockHLC,
        stockOHLC,
        stockVHLC,
        stockVOHLC,
        cylinderColClustered,
        cylinderColStacked,
        cylinderColStacked100,
        cylinderBarClustered,
        cylinderBarStacked,
        cylinderBarStacked100,
        cylinderCol,
        coneColClustered,
        coneColStacked,
        coneColStacked100,
        coneBarClustered,
        coneBarStacked,
        coneBarStacked100,
        coneCol,
        pyramidColClustered,
        pyramidColStacked,
        pyramidColStacked100,
        pyramidBarClustered,
        pyramidBarStacked,
        pyramidBarStacked100,
        pyramidCol,
        line,
        pie,
        xyscatter,
        area,
        doughnut,
        radar,
        histogram,
        boxwhisker,
        pareto,
        regionMap,
        treemap,
        waterfall,
        sunburst,
        funnel,
    }

    enum ChartUnderlineStyle {
        none,
        single,
    }

    enum ChartDisplayBlanksAs {
        notPlotted,
        zero,
        interplotted,
    }

    enum ChartPlotBy {
        rows,
        columns,
    }

    enum ChartSplitType {
        splitByPosition,
        splitByValue,
        splitByPercentValue,
        splitByCustomSplit,
    }

    enum ChartColorScheme {
        colorfulPalette1,
        colorfulPalette2,
        colorfulPalette3,
        colorfulPalette4,
        monochromaticPalette1,
        monochromaticPalette2,
        monochromaticPalette3,
        monochromaticPalette4,
        monochromaticPalette5,
        monochromaticPalette6,
        monochromaticPalette7,
        monochromaticPalette8,
        monochromaticPalette9,
        monochromaticPalette10,
        monochromaticPalette11,
        monochromaticPalette12,
        monochromaticPalette13,
    }

    enum ChartTrendlineType {
        linear,
        exponential,
        logarithmic,
        movingAverage,
        polynomial,
        power,
    }

    enum ShapeZOrder {
        bringToFront,
        bringForward,
        sendToBack,
        sendBackward,
    }

    enum ShapeType {
        unsupported,
        image,
        geometricShape,
        group,
        line,
    }

    enum ShapeScaleType {
        currentSize,
        originalSize,
    }

    enum ShapeScaleFrom {
        scaleFromTopLeft,
        scaleFromMiddle,
        scaleFromBottomRight,
    }

    enum ShapeFillType {
        noFill,
        solid,
        gradient,
        pattern,
        pictureAndTexture,
        mixed,
    }

    enum ShapeFontUnderlineStyle {
        none,
        single,
        double,
        heavy,
        dotted,
        dottedHeavy,
        dash,
        dashHeavy,
        dashLong,
        dashLongHeavy,
        dotDash,
        dotDashHeavy,
        dotDotDash,
        dotDotDashHeavy,
        wavy,
        wavyHeavy,
        wavyDouble,
    }

    enum PictureFormat {
        unknown,
        bmp,
        jpeg,
        gif,
        png,
        svg,
    }

    enum ShapeLineStyle {
        single,
        thickBetweenThin,
        thickThin,
        thinThick,
        thinThin,
    }

    enum ShapeLineDashStyle {
        dash,
        dashDot,
        dashDotDot,
        longDash,
        longDashDot,
        roundDot,
        solid,
        squareDot,
        longDashDotDot,
        systemDash,
        systemDot,
        systemDashDot,
    }

    enum ArrowheadLength {
        short,
        medium,
        long,
    }

    enum ArrowheadStyle {
        none,
        triangle,
        stealth,
        diamond,
        oval,
        open,
    }

    enum ArrowheadWidth {
        narrow,
        medium,
        wide,
    }

    enum BindingType {
        range,
        table,
        text,
    }

    enum BorderIndex {
        edgeTop,
        edgeBottom,
        edgeLeft,
        edgeRight,
        insideVertical,
        insideHorizontal,
        diagonalDown,
        diagonalUp,
    }

    enum BorderLineStyle {
        none,
        continuous,
        dash,
        dashDot,
        dashDotDot,
        dot,
        double,
        slantDashDot,
    }

    enum BorderWeight {
        hairline,
        thin,
        medium,
        thick,
    }

    enum CalculationMode {
        automatic,
        automaticExceptTables,
        manual,
    }

    enum CalculationType {
        recalculate,
        full,
        fullRebuild,
    }

    enum ClearApplyTo {
        all,
        formats,
        contents,
        hyperlinks,
        removeHyperlinks,
    }

    enum ConditionalDataBarAxisFormat {
        automatic,
        none,
        cellMidPoint,
    }

    enum ConditionalDataBarDirection {
        context,
        leftToRight,
        rightToLeft,
    }

    enum ConditionalFormatDirection {
        top,
        bottom,
    }

    enum ConditionalFormatType {
        custom,
        dataBar,
        colorScale,
        iconSet,
        topBottom,
        presetCriteria,
        containsText,
        cellValue,
    }

    enum ConditionalFormatRuleType {
        invalid,
        automatic,
        lowestValue,
        highestValue,
        number,
        percent,
        formula,
        percentile,
    }

    enum ConditionalFormatIconRuleType {
        invalid,
        number,
        percent,
        formula,
        percentile,
    }

    enum ConditionalFormatColorCriterionType {
        invalid,
        lowestValue,
        highestValue,
        number,
        percent,
        formula,
        percentile,
    }

    enum ConditionalTopBottomCriterionType {
        invalid,
        topItems,
        topPercent,
        bottomItems,
        bottomPercent,
    }

    enum ConditionalFormatPresetCriterion {
        invalid,
        blanks,
        nonBlanks,
        errors,
        nonErrors,
        yesterday,
        today,
        tomorrow,
        lastSevenDays,
        lastWeek,
        thisWeek,
        nextWeek,
        lastMonth,
        thisMonth,
        nextMonth,
        aboveAverage,
        belowAverage,
        equalOrAboveAverage,
        equalOrBelowAverage,
        oneStdDevAboveAverage,
        oneStdDevBelowAverage,
        twoStdDevAboveAverage,
        twoStdDevBelowAverage,
        threeStdDevAboveAverage,
        threeStdDevBelowAverage,
        uniqueValues,
        duplicateValues,
    }

    enum ConditionalTextOperator {
        invalid,
        contains,
        notContains,
        beginsWith,
        endsWith,
    }

    enum ConditionalCellValueOperator {
        invalid,
        between,
        notBetween,
        equalTo,
        notEqualTo,
        greaterThan,
        lessThan,
        greaterThanOrEqual,
        lessThanOrEqual,
    }

    enum ConditionalIconCriterionOperator {
        invalid,
        greaterThan,
        greaterThanOrEqual,
    }

    enum ConditionalRangeBorderIndex {
        edgeTop,
        edgeBottom,
        edgeLeft,
        edgeRight,
    }

    enum ConditionalRangeBorderLineStyle {
        none,
        continuous,
        dash,
        dashDot,
        dashDotDot,
        dot,
    }

    enum ConditionalRangeFontUnderlineStyle {
        none,
        single,
        double,
    }

    enum DataValidationType {
        none,
        wholeNumber,
        decimal,
        list,
        date,
        time,
        textLength,
        custom,
        inconsistent,
        mixedCriteria,
    }

    enum DataValidationOperator {
        between,
        notBetween,
        equalTo,
        notEqualTo,
        greaterThan,
        lessThan,
        greaterThanOrEqualTo,
        lessThanOrEqualTo,
    }

    enum DataValidationAlertStyle {
        stop,
        warning,
        information,
    }

    enum DeleteShiftDirection {
        up,
        left,
    }

    enum DynamicFilterCriteria {
        unknown,
        aboveAverage,
        allDatesInPeriodApril,
        allDatesInPeriodAugust,
        allDatesInPeriodDecember,
        allDatesInPeriodFebruray,
        allDatesInPeriodJanuary,
        allDatesInPeriodJuly,
        allDatesInPeriodJune,
        allDatesInPeriodMarch,
        allDatesInPeriodMay,
        allDatesInPeriodNovember,
        allDatesInPeriodOctober,
        allDatesInPeriodQuarter1,
        allDatesInPeriodQuarter2,
        allDatesInPeriodQuarter3,
        allDatesInPeriodQuarter4,
        allDatesInPeriodSeptember,
        belowAverage,
        lastMonth,
        lastQuarter,
        lastWeek,
        lastYear,
        nextMonth,
        nextQuarter,
        nextWeek,
        nextYear,
        thisMonth,
        thisQuarter,
        thisWeek,
        thisYear,
        today,
        tomorrow,
        yearToDate,
        yesterday,
    }

    enum FilterDatetimeSpecificity {
        year,
        month,
        day,
        hour,
        minute,
        second,
    }

    enum FilterOn {
        bottomItems,
        bottomPercent,
        cellColor,
        dynamic,
        fontColor,
        values,
        topItems,
        topPercent,
        icon,
        custom,
    }

    enum FilterOperator {
        and,
        or,
    }

    enum HorizontalAlignment {
        general,
        left,
        center,
        right,
        fill,
        justify,
        centerAcrossSelection,
        distributed,
    }

    enum IconSet {
        invalid,
        threeArrows,
        threeArrowsGray,
        threeFlags,
        threeTrafficLights1,
        threeTrafficLights2,
        threeSigns,
        threeSymbols,
        threeSymbols2,
        fourArrows,
        fourArrowsGray,
        fourRedToBlack,
        fourRating,
        fourTrafficLights,
        fiveArrows,
        fiveArrowsGray,
        fiveRating,
        fiveQuarters,
        threeStars,
        threeTriangles,
        fiveBoxes,
    }

    enum ImageFittingMode {
        fit,
        fitAndCenter,
        fill,
    }

    enum InsertShiftDirection {
        down,
        right,
    }

    enum NamedItemScope {
        worksheet,
        workbook,
    }

    enum NamedItemType {
        string,
        integer,
        double,
        boolean,
        range,
        error,
        array,
    }

    enum RangeUnderlineStyle {
        none,
        single,
        double,
        singleAccountant,
        doubleAccountant,
    }

    enum SheetVisibility {
        visible,
        hidden,
        veryHidden,
    }

    enum RangeValueType {
        unknown,
        empty,
        string,
        integer,
        double,
        boolean,
        error,
        richValue,
    }

    enum SearchDirection {
        forward,
        backwards,
    }

    enum SortOrientation {
        rows,
        columns,
    }

    enum SortOn {
        value,
        cellColor,
        fontColor,
        icon,
    }

    enum SortDataOption {
        normal,
        textAsNumber,
    }

    enum SortMethod {
        pinYin,
        strokeCount,
    }

    enum VerticalAlignment {
        top,
        center,
        bottom,
        justify,
        distributed,
    }

    enum DocumentPropertyType {
        number,
        boolean,
        date,
        string,
        float,
    }

    enum SubtotalLocationType {
        atTop,
        atBottom,
        off,
    }

    enum PivotLayoutType {
        compact,
        tabular,
        outline,
    }

    enum ProtectionSelectionMode {
        normal,
        unlocked,
        none,
    }

    enum PageOrientation {
        portrait,
        landscape,
    }

    enum PaperType {
        letter,
        letterSmall,
        tabloid,
        ledger,
        legal,
        statement,
        executive,
        a3,
        a4,
        a4Small,
        a5,
        b4,
        b5,
        folio,
        quatro,
        paper10x14,
        paper11x17,
        note,
        envelope9,
        envelope10,
        envelope11,
        envelope12,
        envelope14,
        csheet,
        dsheet,
        esheet,
        envelopeDL,
        envelopeC5,
        envelopeC3,
        envelopeC4,
        envelopeC6,
        envelopeC65,
        envelopeB4,
        envelopeB5,
        envelopeB6,
        envelopeItaly,
        envelopeMonarch,
        envelopePersonal,
        fanfoldUS,
        fanfoldStdGerman,
        fanfoldLegalGerman,
    }

    enum ReadingOrder {
        context,
        leftToRight,
        rightToLeft,
    }

    enum BuiltInStyle {
        normal,
        comma,
        currency,
        percent,
        wholeComma,
        wholeDollar,
        hlink,
        hlinkTrav,
        note,
        warningText,
        emphasis1,
        emphasis2,
        emphasis3,
        sheetTitle,
        heading1,
        heading2,
        heading3,
        heading4,
        input,
        output,
        calculation,
        checkCell,
        linkedCell,
        total,
        good,
        bad,
        neutral,
        accent1,
        accent1_20,
        accent1_40,
        accent1_60,
        accent2,
        accent2_20,
        accent2_40,
        accent2_60,
        accent3,
        accent3_20,
        accent3_40,
        accent3_60,
        accent4,
        accent4_20,
        accent4_40,
        accent4_60,
        accent5,
        accent5_20,
        accent5_40,
        accent5_60,
        accent6,
        accent6_20,
        accent6_40,
        accent6_60,
        explanatoryText,
    }

    enum PrintErrorType {
        asDisplayed,
        blank,
        dash,
        notAvailable,
    }

    enum WorksheetPositionType {
        none,
        before,
        after,
        beginning,
        end,
    }

    enum PrintComments {
        noComments,
        endSheet,
        inPlace,
    }

    enum PrintOrder {
        downThenOver,
        overThenDown,
    }

    enum PrintMarginUnit {
        points,
        inches,
        centimeters,
    }

    enum HeaderFooterState {
        default,
        firstAndDefault,
        oddAndEven,
        firstOddAndEven,
    }

    enum AutoFillType {
        fillDefault,
        fillCopy,
        fillSeries,
        fillFormats,
        fillValues,
        fillDays,
        fillWeekdays,
        fillMonths,
        fillYears,
        linearTrend,
        growthTrend,
        flashFill,
    }

    enum GroupOption {
        byRows,
        byColumns,
    }

    enum RangeCopyType {
        all,
        formulas,
        values,
        formats,
    }

    enum LinkedDataTypeState {
        none,
        validLinkedData,
        disambiguationNeeded,
        brokenLinkedData,
        fetchingData,
    }

    enum GeometricShapeType {
        lineInverse,
        triangle,
        rightTriangle,
        rectangle,
        diamond,
        parallelogram,
        trapezoid,
        nonIsoscelesTrapezoid,
        pentagon,
        hexagon,
        heptagon,
        octagon,
        decagon,
        dodecagon,
        star4,
        star5,
        star6,
        star7,
        star8,
        star10,
        star12,
        star16,
        star24,
        star32,
        roundRectangle,
        round1Rectangle,
        round2SameRectangle,
        round2DiagonalRectangle,
        snipRoundRectangle,
        snip1Rectangle,
        snip2SameRectangle,
        snip2DiagonalRectangle,
        plaque,
        ellipse,
        teardrop,
        homePlate,
        chevron,
        pieWedge,
        pie,
        blockArc,
        donut,
        noSmoking,
        rightArrow,
        leftArrow,
        upArrow,
        downArrow,
        stripedRightArrow,
        notchedRightArrow,
        bentUpArrow,
        leftRightArrow,
        upDownArrow,
        leftUpArrow,
        leftRightUpArrow,
        quadArrow,
        leftArrowCallout,
        rightArrowCallout,
        upArrowCallout,
        downArrowCallout,
        leftRightArrowCallout,
        upDownArrowCallout,
        quadArrowCallout,
        bentArrow,
        uturnArrow,
        circularArrow,
        leftCircularArrow,
        leftRightCircularArrow,
        curvedRightArrow,
        curvedLeftArrow,
        curvedUpArrow,
        curvedDownArrow,
        swooshArrow,
        cube,
        can,
        lightningBolt,
        heart,
        sun,
        moon,
        smileyFace,
        irregularSeal1,
        irregularSeal2,
        foldedCorner,
        bevel,
        frame,
        halfFrame,
        corner,
        diagonalStripe,
        chord,
        arc,
        leftBracket,
        rightBracket,
        leftBrace,
        rightBrace,
        bracketPair,
        bracePair,
        callout1,
        callout2,
        callout3,
        accentCallout1,
        accentCallout2,
        accentCallout3,
        borderCallout1,
        borderCallout2,
        borderCallout3,
        accentBorderCallout1,
        accentBorderCallout2,
        accentBorderCallout3,
        wedgeRectCallout,
        wedgeRRectCallout,
        wedgeEllipseCallout,
        cloudCallout,
        cloud,
        ribbon,
        ribbon2,
        ellipseRibbon,
        ellipseRibbon2,
        leftRightRibbon,
        verticalScroll,
        horizontalScroll,
        wave,
        doubleWave,
        plus,
        flowChartProcess,
        flowChartDecision,
        flowChartInputOutput,
        flowChartPredefinedProcess,
        flowChartInternalStorage,
        flowChartDocument,
        flowChartMultidocument,
        flowChartTerminator,
        flowChartPreparation,
        flowChartManualInput,
        flowChartManualOperation,
        flowChartConnector,
        flowChartPunchedCard,
        flowChartPunchedTape,
        flowChartSummingJunction,
        flowChartOr,
        flowChartCollate,
        flowChartSort,
        flowChartExtract,
        flowChartMerge,
        flowChartOfflineStorage,
        flowChartOnlineStorage,
        flowChartMagneticTape,
        flowChartMagneticDisk,
        flowChartMagneticDrum,
        flowChartDisplay,
        flowChartDelay,
        flowChartAlternateProcess,
        flowChartOffpageConnector,
        actionButtonBlank,
        actionButtonHome,
        actionButtonHelp,
        actionButtonInformation,
        actionButtonForwardNext,
        actionButtonBackPrevious,
        actionButtonEnd,
        actionButtonBeginning,
        actionButtonReturn,
        actionButtonDocument,
        actionButtonSound,
        actionButtonMovie,
        gear6,
        gear9,
        funnel,
        mathPlus,
        mathMinus,
        mathMultiply,
        mathDivide,
        mathEqual,
        mathNotEqual,
        cornerTabs,
        squareTabs,
        plaqueTabs,
        chartX,
        chartStar,
        chartPlus,
    }

    enum ConnectorType {
        straight,
        elbow,
        curve,
    }

    enum ContentType {
        plain,
        mention,
    }

    enum SpecialCellType {
        conditionalFormats,
        dataValidations,
        blanks,
        constants,
        formulas,
        sameConditionalFormat,
        sameDataValidation,
        visible,
    }

    enum SpecialCellValueType {
        all,
        errors,
        errorsLogical,
        errorsNumbers,
        errorsText,
        errorsLogicalNumber,
        errorsLogicalText,
        errorsNumberText,
        logical,
        logicalNumbers,
        logicalText,
        logicalNumbersText,
        numbers,
        numbersText,
        text,
    }

    enum Placement {
        twoCell,
        oneCell,
        absolute,
    }

    enum FillPattern {
        none,
        solid,
        gray50,
        gray75,
        gray25,
        horizontal,
        vertical,
        down,
        up,
        checker,
        semiGray75,
        lightHorizontal,
        lightVertical,
        lightDown,
        lightUp,
        grid,
        crissCross,
        gray16,
        gray8,
        linearGradient,
        rectangularGradient,
    }

    enum ShapeTextHorizontalAlignment {
        left,
        center,
        right,
        justify,
        justifyLow,
        distributed,
        thaiDistributed,
    }

    enum ShapeTextVerticalAlignment {
        top,
        middle,
        bottom,
        justified,
        distributed,
    }

    enum ShapeTextVerticalOverflow {
        overflow,
        ellipsis,
        clip,
    }

    enum ShapeTextHorizontalOverflow {
        overflow,
        clip,
    }

    enum ShapeTextReadingOrder {
        leftToRight,
        rightToLeft,
    }

    enum ShapeTextOrientation {
        horizontal,
        vertical,
        vertical270,
        wordArtVertical,
        eastAsianVertical,
        mongolianVertical,
        wordArtVerticalRTL,
    }

    enum ShapeAutoSize {
        autoSizeNone,
        autoSizeTextToFitShape,
        autoSizeShapeToFitText,
        autoSizeMixed,
    }

    enum SlicerSortType {
        dataSourceOrder,
        ascending,
        descending,
    }
}
