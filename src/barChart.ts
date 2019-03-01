
module powerbi.extensibility.visual {
    // powerbi.visuals
    import ISelectionId = powerbi.visuals.ISelectionId;

    /**
     * Interface for BarCharts viewmodel.
     *
     * @interface
     * @property {BarChartDataPoint[]} dataPoints - Set of data points the visual will render.
     * @property {number} dataMax                 - Maximum data value in the set of data points.
     */
    interface BarChartViewModel {
        dataPoints: BarChartDataPoint[];
        dataMax: number;
        settings: BarChartSettings;
    };

    /**
     * Interface for BarChart data points.
     *
     * @interface
     * @property {number} value             - Data value for point.
     * @property {string} category          - Corresponding category of data value.
     * @property {string} color             - Color corresponding to data point.
     * @property {string} endColor          - Gradient end color corresponding to data point.
     * @property {ISelectionId} selectionId - Id assigned to data point for cross filtering
     *                                        and visual interaction.
     */
    interface BarChartDataPoint {
        value: PrimitiveValue;
        category: string;
        color: string;
        endColor: string;
        strokeColor: string;
        strokeWidth: number;
        selectionId: ISelectionId;
    };

    /**
     * Interface for BarChart settings.
     *
     * @interface
     * @property {{show:boolean}} enableAxis - Object property that allows axis to be enabled.
     * @property {{generalView.opacity:number}} Bars Opacity - Controls opacity of plotted bars, values range between 10 (almost transparent) to 100 (fully opaque, default)
     * @property {{generalView.showHelpLink:boolean}} Show Help Button - When TRUE, the plot displays a button which launch a link to documentation.
     */
    interface BarChartSettings {
        enableAxis: {
            show: boolean;
            fill: string;
        };

        generalView: {
            opacity: number;
            showHelpLink: boolean;
            helpLinkColor: string;
        };
    }

    /**
     * Function that converts queried data into a view model that will be used by the visual.
     *
     * @function
     * @param {VisualUpdateOptions} options - Contains references to the size of the container
     *                                        and the dataView which contains all the data
     *                                        the visual had queried.
     * @param {IVisualHost} host            - Contains references to the host which contains services
     */
    function visualTransform(options: VisualUpdateOptions, host: IVisualHost): BarChartViewModel {
        let dataViews = options.dataViews;
        let defaultSettings: BarChartSettings = {
            enableAxis: {
                show: false,
                fill: "#000000",
            },
            generalView: {
                opacity: 100,
                showHelpLink: false,
                helpLinkColor: "#80B0E0",
            }
        };
        let viewModel: BarChartViewModel = {
            dataPoints: [],
            dataMax: 0,
            settings: <BarChartSettings>{}
        };

        if (!dataViews
            || !dataViews[0]
            || !dataViews[0].categorical
            || !dataViews[0].categorical.categories
            || !dataViews[0].categorical.categories[0].source
            || !dataViews[0].categorical.values
        ) {
            return viewModel;
        }

        let categorical = dataViews[0].categorical;
        let category = categorical.categories[0];
        let dataValue = categorical.values[0];

        let barChartDataPoints: BarChartDataPoint[] = [];
        let dataMax: number;

        let colorPalette: ISandboxExtendedColorPalette = host.colorPalette;
        let objects = dataViews[0].metadata.objects;

        const strokeColor: string = getColumnStrokeColor(colorPalette);

        let barChartSettings: BarChartSettings = {
            enableAxis: {
                show: getValue<boolean>(objects, 'enableAxis', 'show', defaultSettings.enableAxis.show),
                fill: getAxisTextFillColor(objects, colorPalette, defaultSettings.enableAxis.fill),
            },
            generalView: {
                opacity: getValue<number>(objects, 'generalView', 'opacity', defaultSettings.generalView.opacity),
                showHelpLink: getValue<boolean>(objects, 'generalView', 'showHelpLink', defaultSettings.generalView.showHelpLink),
                helpLinkColor: strokeColor,
            },
        };

        const strokeWidth: number = getColumnStrokeWidth(colorPalette.isHighContrast);

        for (let i = 0, len = Math.max(category.values.length, dataValue.values.length); i < len; i++) {
            const color: string = getColumnColorByIndex(category, i, colorPalette, false);
            const endColor: string = getColumnColorByIndex(category, i, colorPalette, true);

            const selectionId: ISelectionId = host.createSelectionIdBuilder()
                .withCategory(category, i)
                .createSelectionId();

            barChartDataPoints.push({
                color,
                endColor,
                strokeColor,
                strokeWidth,
                selectionId,
                value: dataValue.values[i],
                category: `${category.values[i]}`,
            });
        }

        dataMax = <number>dataValue.maxLocal;

        return {
            dataPoints: barChartDataPoints,
            dataMax: dataMax,
            settings: barChartSettings,
        };
    }

    function shadeColor(color: string, percent: number) {
        let R = parseInt(color.substring(1, 3), 16);
        let G = parseInt(color.substring(3, 5), 16);
        let B = parseInt(color.substring(5, 7), 16);

        R = parseInt(`${R * (100 + percent) / 100}`);
        G = parseInt(`${G * (100 + percent) / 100}`);
        B = parseInt(`${B * (100 + percent) / 100}`);

        R = (R < 255) ? R : 255;
        G = (G < 255) ? G : 255;
        B = (B < 255) ? B : 255;

        const RR = ((R.toString(16).length === 1) ? "0" + R.toString(16) : R.toString(16));
        const GG = ((G.toString(16).length === 1) ? "0" + G.toString(16) : G.toString(16));
        const BB = ((B.toString(16).length === 1) ? "0" + B.toString(16) : B.toString(16));
        return "#" + RR + GG + BB;
    }

    function getColumnColorByIndex(
        category: DataViewCategoryColumn,
        index: number,
        colorPalette: ISandboxExtendedColorPalette,
        isEndColor: boolean
    ): string {
        if (colorPalette.isHighContrast) {
            return colorPalette.background.value;
        }

        const categoryColor = colorPalette.getColor(`${category.values[index]}`).value;
        const categoryEndColor = shadeColor(categoryColor, -40);
        console.log('color', categoryColor, categoryEndColor);
        const defaultColor: Fill = {
            solid: {
                color: isEndColor ? categoryEndColor : categoryColor,
            }
        };

        return getCategoricalObjectValue<Fill>(
            category,
            index,
            'colorSelector',
            isEndColor ? 'endFill' : 'fill',
            defaultColor
        ).solid.color;
    }

    function getColumnStrokeColor(colorPalette: ISandboxExtendedColorPalette): string {
        return colorPalette.isHighContrast
            ? colorPalette.foreground.value
            : null;
    }

    function getColumnStrokeWidth(isHighContrast: boolean): number {
        return isHighContrast
            ? 2
            : 0;
    }

    function getAxisTextFillColor(
        objects: DataViewObjects,
        colorPalette: ISandboxExtendedColorPalette,
        defaultColor: string
    ): string {
        if (colorPalette.isHighContrast) {
            return colorPalette.foreground.value;
        }

        return getValue<Fill>(
            objects,
            "enableAxis",
            "fill",
            {
                solid: {
                    color: defaultColor,
                }
            },
        ).solid.color;
    }

    export class BarChart implements IVisual {
        private svg: d3.Selection<SVGElement>;
        private host: IVisualHost;
        private selectionManager: ISelectionManager;
        private barContainer: d3.Selection<SVGElement>;
        private xAxis: d3.Selection<SVGElement>;
        private barDataPoints: BarChartDataPoint[];
        private barChartSettings: BarChartSettings;
        private tooltipServiceWrapper: ITooltipServiceWrapper;
        private locale: string;
        private helpLinkElement: d3.Selection<any>;
        private element: HTMLElement;
        private isLandingPageOn: boolean;
        private LandingPageRemoved: boolean;
        private LandingPage: d3.Selection<any>;

        private barSelection: d3.selection.Update<BarChartDataPoint>;

        static Config = {
            xScalePadding: .7,
            solidOpacity: 1,
            transparentOpacity: 0.4,
            margins: {
                top: 0,
                right: 0,
                bottom: 25,
                left: 30,
            },
            xAxisFontMultiplier: 0.04,
        };

        /**
         * Creates instance of BarChart. This method is only called once.
         *
         * @constructor
         * @param {VisualConstructorOptions} options - Contains references to the element that will
         *                                             contain the visual and a reference to the host
         *                                             which contains services.
         */
        constructor(options: VisualConstructorOptions) {
            this.host = options.host;
            this.element = options.element;
            this.selectionManager = options.host.createSelectionManager();

            this.selectionManager.registerOnSelectCallback(() => {
                this.syncSelectionState(this.barSelection, this.selectionManager.getSelectionIds() as ISelectionId[]);
            });

            this.tooltipServiceWrapper = createTooltipServiceWrapper(this.host.tooltipService, options.element);

            this.svg = d3.select(options.element)
                .append('svg')
                .classed('barChart', true);

            this.locale = options.host.locale;

            this.barContainer = this.svg
                .append('g')
                .classed('barContainer', true);

            this.xAxis = this.svg
                .append('g')
                .classed('xAxis', true)
                .style("fill", "url(#linearGradient)");

            const helpLinkElement: Element = this.createHelpLinkElement();
            options.element.appendChild(helpLinkElement);

            this.helpLinkElement = d3.select(helpLinkElement);
        }

        /**
         * Updates the state of the visual. Every sequential databinding and resize will call update.
         *
         * @function
         * @param {VisualUpdateOptions} options - Contains references to the size of the container
         *                                        and the dataView which contains all the data
         *                                        the visual had queried.
         */
        public update(options: VisualUpdateOptions) {
            let viewModel: BarChartViewModel = visualTransform(options, this.host);
            let settings = this.barChartSettings = viewModel.settings;
            this.barDataPoints = viewModel.dataPoints;

            // Turn on landing page in capabilities and remove comment to turn on landing page!
            // this.HandleLandingPage(options);

            let width = options.viewport.width;
            let height = options.viewport.height;

            this.svg.attr({
                width: width,
                height: height
            });

            if (settings.enableAxis.show) {
                let margins = BarChart.Config.margins;
                height -= margins.bottom;
            }

            this.helpLinkElement
                .classed("hidden", !settings.generalView.showHelpLink)
                .style({
                    "border-color": settings.generalView.helpLinkColor,
                    "color": settings.generalView.helpLinkColor,
                });

            this.xAxis.style({
                "font-size": d3.min([height, width]) * BarChart.Config.xAxisFontMultiplier,
                "fill": settings.enableAxis.fill,
            });

            let yScale = d3.scale.linear()
                .domain([0, viewModel.dataMax])
                .range([height, 0]);

            let xScale = d3.scale.ordinal()
                .domain(viewModel.dataPoints.map(d => d.category))
                .rangeRoundBands([0, width], BarChart.Config.xScalePadding, 0.2);

            let xAxis = d3.svg.axis()
                .scale(xScale)
                .orient('bottom');

            this.xAxis.attr('transform', 'translate(0, ' + height + ')')
                .call(xAxis);

            this.barSelection = this.barContainer
                .selectAll('.bar')
                .data(this.barDataPoints);

            this.barSelection
                .enter()
                .append('rect')
                .classed('bar', true);

            const opacity: number = viewModel.settings.generalView.opacity / 100;

            // create gradients
            this.barDataPoints.forEach(dataPoint => {
                const gradientId = `gradient_${dataPoint.color.substr(1)}_${dataPoint.endColor.substr(1)}`;
                // if gradient doesn't exist yet create it
                if (this.svg.select(`#${gradientId}`).empty()) {
                    const linearGradient = this.svg.append("defs")
                        .append("linearGradient")
                        .attr("id", gradientId)
                        .attr("x1", "0%").attr("y1", "0%")
                        .attr("x2", "0%").attr("y2", "100%");
                    linearGradient.append("stop")
                        .attr("offset", "0%")
                        .attr("stop-color", dataPoint.color)
                        .attr("stop-opacity", 1);
                    linearGradient.append("stop")
                        .attr("offset", "100%")
                        .attr("stop-color", dataPoint.endColor)
                        .attr("stop-opacity", 1);
                }
            });

            this.barSelection
                .attr({
                    width: xScale.rangeBand(),
                    height: d => height - yScale(<number>d.value),
                    y: d => yScale(<number>d.value),
                    x: d => xScale(d.category),
                })
                .style({
                    'fill-opacity': opacity,
                    'stroke-opacity': opacity,
                    fill: (dataPoint: BarChartDataPoint) => `url(#gradient_${dataPoint.color.substr(1)}_${dataPoint.endColor.substr(1)})`,
                    stroke: (dataPoint: BarChartDataPoint) => dataPoint.strokeColor,
                    "stroke-width": (dataPoint: BarChartDataPoint) => `${dataPoint.strokeWidth}px`,
                });

            this.tooltipServiceWrapper.addTooltip(this.barContainer.selectAll('.bar'),
                (tooltipEvent: TooltipEventArgs<BarChartDataPoint>) => this.getTooltipData(tooltipEvent.data),
                (tooltipEvent: TooltipEventArgs<BarChartDataPoint>) => tooltipEvent.data.selectionId
            );

            this.syncSelectionState(
                this.barSelection,
                this.selectionManager.getSelectionIds() as ISelectionId[]
            );

            this.barSelection.on('click', (d) => {
                // Allow selection only if the visual is rendered in a view that supports interactivity (e.g. Report)
                if (this.host.allowInteractions) {
                    const isCtrlPressed: boolean = (d3.event as MouseEvent).ctrlKey;

                    this.selectionManager
                        .select(d.selectionId, isCtrlPressed)
                        .then((ids: ISelectionId[]) => {
                            this.syncSelectionState(this.barSelection, ids);
                        });

                    (<Event>d3.event).stopPropagation();
                }
            });

            this.barSelection
                .exit()
                .remove();

            // Clear selection when clicking outside a bar
            this.svg.on('click', (d) => {
                if (this.host.allowInteractions) {
                    this.selectionManager
                        .clear()
                        .then(() => {
                            this.syncSelectionState(this.barSelection, []);
                        });
                }
            });
            // handle context menu
            this.svg.on('contextmenu', () => {
                const mouseEvent: MouseEvent = d3.event as MouseEvent;
                const eventTarget: EventTarget = mouseEvent.target;
                let dataPoint = d3.select(eventTarget).datum();
                this.selectionManager.showContextMenu(dataPoint ? dataPoint.selectionId : {}, {
                    x: mouseEvent.clientX,
                    y: mouseEvent.clientY
                });
                mouseEvent.preventDefault();
            });
        }

        private syncSelectionState(
            selection: d3.Selection<BarChartDataPoint>,
            selectionIds: ISelectionId[]
        ): void {
            if (!selection || !selectionIds) {
                return;
            }

            if (!selectionIds.length) {
                selection.style({
                    "fill-opacity": null,
                    "stroke-opacity": null,
                });

                return;
            }

            const self: this = this;

            selection.each(function (barDataPoint: BarChartDataPoint) {
                const isSelected: boolean = self.isSelectionIdInArray(selectionIds, barDataPoint.selectionId);

                const opacity: number = isSelected
                    ? BarChart.Config.solidOpacity
                    : BarChart.Config.transparentOpacity;

                d3.select(this).style({
                    "fill-opacity": opacity,
                    "stroke-opacity": opacity,
                });
            });
        }

        private isSelectionIdInArray(selectionIds: ISelectionId[], selectionId: ISelectionId): boolean {
            if (!selectionIds || !selectionId) {
                return false;
            }

            return selectionIds.some((currentSelectionId: ISelectionId) => {
                return currentSelectionId.includes(selectionId);
            });
        }

        /**
         * Enumerates through the objects defined in the capabilities and adds the properties to the format pane
         *
         * @function
         * @param {EnumerateVisualObjectInstancesOptions} options - Map of defined objects
         */
        public enumerateObjectInstances(options: EnumerateVisualObjectInstancesOptions): VisualObjectInstanceEnumeration {
            let objectName = options.objectName;
            let objectEnumeration: VisualObjectInstance[] = [];

            switch (objectName) {
                case 'enableAxis':
                    objectEnumeration.push({
                        objectName: objectName,
                        properties: {
                            show: this.barChartSettings.enableAxis.show,
                            fill: this.barChartSettings.enableAxis.fill,
                        },
                        selector: null
                    });
                    break;
                case 'colorSelector':
                    for (let barDataPoint of this.barDataPoints) {
                        objectEnumeration.push({
                            objectName: objectName,
                            displayName: barDataPoint.category,
                            properties: {
                                fill: {
                                    solid: {
                                        color: barDataPoint.color
                                    }
                                },
                                endFill: {
                                    solid: {
                                        color: barDataPoint.endColor
                                    }
                                }
                            },
                            selector: barDataPoint.selectionId.getSelector()
                        });
                    }
                    break;
                case 'generalView':
                    objectEnumeration.push({
                        objectName: objectName,
                        properties: {
                            opacity: this.barChartSettings.generalView.opacity,
                            showHelpLink: this.barChartSettings.generalView.showHelpLink
                        },
                        validValues: {
                            opacity: {
                                numberRange: {
                                    min: 10,
                                    max: 100
                                }
                            }
                        },
                        selector: null
                    });
                    break;
            };

            return objectEnumeration;
        }

        /**
         * Destroy runs when the visual is removed. Any cleanup that the visual needs to
         * do should be done here.
         *
         * @function
         */
        public destroy(): void {
            // Perform any cleanup tasks here
        }

        private getTooltipData(value: any): VisualTooltipDataItem[] {
            let language = getLocalizedString(this.locale, "LanguageKey");
            return [{
                displayName: value.category,
                value: value.value.toString(),
                color: value.color,
                header: language && "displayed language " + language
            }];
        }

        private createHelpLinkElement(): Element {
            let linkElement = document.createElement("a");
            linkElement.textContent = "?";
            linkElement.setAttribute("title", "Open documentation");
            linkElement.setAttribute("class", "helpLink");
            linkElement.addEventListener("click", () => {
                this.host.launchUrl("https://microsoft.github.io/PowerBI-visuals/tutorials/building-bar-chart/adding-url-launcher-element-to-the-bar-chart/");
            });
            return linkElement;
        };

        private HandleLandingPage(options: VisualUpdateOptions) {
            if (!options.dataViews || !options.dataViews.length) {
                if (!this.isLandingPageOn) {
                    this.isLandingPageOn = true;
                    const SampleLandingPage: Element = this.createSampleLandingPage();
                    this.element.appendChild(SampleLandingPage);

                    this.LandingPage = d3.select(SampleLandingPage);
                }

            } else {
                    if (this.isLandingPageOn && !this.LandingPageRemoved) {
                        this.LandingPageRemoved = true;
                        this.LandingPage.remove();
                }
            }
        }

        private createSampleLandingPage(): Element {
            let div = document.createElement("div");

            let header = document.createElement("h1");
            header.textContent = "Sample Bar Chart Landing Page";
            header.setAttribute("class", "LandingPage");
            let p1 = document.createElement("a");
            p1.setAttribute("class", "LandingPageHelpLink");
            p1.textContent = "Learn more about Landing page";

            p1.addEventListener("click", () => {
                this.host.launchUrl("https://microsoft.github.io/PowerBI-visuals/docs/overview/");
            });

            div.appendChild(header);
            div.appendChild(p1);

            return div;
        }
    }
}
