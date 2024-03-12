%let pgm=utl-sas-ods-excel-to-create-excel-report-and-separate-png-graph-finally-r-for-layout-in-excel;

sas ods excel to create excel report and separate png graph finally r for layout in excel

github
https://tinyurl.com/2js68yun
https://github.com/rogerjdeangelis/utl-sas-ods-excel-to-create-excel-report-and-separate-png-graph-finally-r-for-layout-in-excel

PROCESS

   Create workbook and sheet with proc report table
   Create a separate sqlot graph
   Use R to position plot beside excel report

Related repos on end

/*               _     _
 _ __  _ __ ___ | |__ | | ___ _ __ ___
| `_ \| `__/ _ \| `_ \| |/ _ \ `_ ` _ \
| |_) | | | (_) | |_) | |  __/ | | | | |
| .__/|_|  \___/|_.__/|_|\___|_| |_| |_|
|_|
*/

/**************************************************************************************************************************/
/*                           |                                        |                                                   */
/*                           |                                        |                                                   */
/*            INPUT          |           PROCESS                      |                OUTPUT(d:/x;s/sbys.xlsx)           */
/*                           |                                        |                                                   */
/* SASHELP.PRDSALE obs=960   | * CREATE RXCEL SHEET WITH REPORT       | EXCEL   A       B        C         PRODUCTS       */
/*                           |                                        |  ROW -------|--------|--------                    */
/* Ob ACTUAL COUNTRY PRODUCT | %utlfkil(d:\xls\sbys.xlsx);            |                                  ---+--+--+--+--  */
/*                           | %utlfkil(d:\png\sbs.png);              |  1   Country  Product  Actual   6+             +6 */
/*  1   925  CANADA   SOFA   |                                        |                                  |             |  */
/*  2   999  CANADA   SOFA   | ods excel file="d:\xls\sbys.xlsx";     |  2   CANADA   BED      $47,729  4+       ** ** +4 */
/*  3   608  CANADA   SOFA   | ods excel options                      |  3            CHAIR    $50,239   |       ** ** |  */
/* ....                      |   (sheet_name="sheet1" start_at="A1"); |  4            DESK     $52,187  2+ ** ** ** ** +2 */
/* 958   884  GERMANY  DESK  |                                        |  5            SOFA     $50,135   | ** ** ** ** |  */
/* 959   689  GERMANY  DESK  | proc report data=sashelp.prdsale       |  6            TABLE    $46,700   ---+--+--+--+--  */
/* 960   646  GERMANY  DESK  | (keep=country product actual where=    |  7   GERMANY  BED      $46,134      B  C  D  T    */
/*                           | (country in ("CANADA","GERMANY")));    |  8           CHAIR     $47,105      E  H  E  A    */
/*                           | column country product actual;         |  9            DESK     $48,502      D  A  S  B    */
/*                           | define country / group;                |                                        I  K  L    */
/*                           | define product / group;                |  ------                                R          */
/*                           | rbreak after / summarize;              |  SHEET1                                           */
/*                           | run;quit;                              |  ------                                           */
/*                           |                                        |                                                   */
/*                           | Ods excel close;                       |                                                   */
/*                           |                                        |                                                   */
/*                           | * CREATE PNG FILE WITH HISTOGRAM;      |                                                   */
/*                           |                                        |                                                   */
/*                           | ods listing style=journal;             |                                                   */
/*                           | ods listing  gpath='d:/png';           |                                                   */
/*                           | ods graphics on / width=4in            |                                                   */
/*                           |  imagefmt=png imagename="sbys";        |                                                   */
/*                           |                                        |                                                   */
/*                           | proc sgplot data=sashelp.class;        |                                                   */
/*                           | format age 2.;                         |                                                   */
/*                           | vbar age /datalabel;                   |                                                   */
/*                           | run;quit;                              |                                                   */
/*                           |                                        |                                                   */
/*                           | ods graphics off;                      |                                                   */
/*                           |                                        |                                                   */
/*                           | * ADD HISTOGRAM RIGHT OF REPORT;       |                                                   */
/*                           |                                        |                                                   */
/*                           | %utl_submit_r64('                      |                                                   */
/*                           | library(openxlsx);                     |                                                   */
/*                           | wb<-loadWorkbook(                      |                                                   */
/*                           |     file="d:/xls/sbys.xlsx");          |                                                   */
/*                           | insertImage(                           |                                                   */
/*                           |   wb                                   |                                                   */
/*                           |   ,"sheet1"                            |                                                   */
/*                           |   ,"d:/png/sbys.png"                   |                                                   */
/*                           |   ,width    = 8                        |                                                   */
/*                           |   ,height   = 6                        |                                                   */
/*                           |   ,startRow = 1                        |                                                   */
/*                           |   ,startCol = "E"                      |                                                   */
/*                           |   ,units = "cm");                      |                                                   */
/*                           | saveWorkbook(                          |                                                   */
/*                           |     wb                                 |                                                   */
/*                           |   ,"d:/xls/sbys.xlsx"                  |                                                   */
/*                           |   ,overwrite = TRUE);                  |                                                   */
/*                           | ');                                    |                                                   */
/*                           |                                        |                                                   */
/*****************************|********************************************************************************************/

/*                   _
(_)_ __  _ __  _   _| |_
| | `_ \| `_ \| | | | __|
| | | | | |_) | |_| | |_
|_|_| |_| .__/ \__,_|\__|
        |_|
*/

SASHELP.PRDSALE

sashelp.prdsale
(keep=country product actual where=
(country in ("CANADA","GERMANY")));

/*
 _ __  _ __ ___   ___ ___  ___ ___
| `_ \| `__/ _ \ / __/ _ \/ __/ __|
| |_) | | | (_) | (_|  __/\__ \__ \
| .__/|_|  \___/ \___\___||___/___/
|_|
*/

 * CREATE RXCEL SHEET WITH REPORT

 %utlfkil(d:\xls\sbys.xlsx);
 %utlfkil(d:\png\sbs.png);

 ods excel file="d:\xls\sbys.xlsx";
 ods excel options
   (sheet_name="sheet1" start_at="A1");

 proc report data=sashelp.prdsale
 (keep=country product actual where=
 (country in ("CANADA","GERMANY")));
 column country product actual;
 define country / group;
 define product / group;
 rbreak after / summarize;
 run;quit;

 Ods excel close;

 * CREATE PNG FILE WITH HISTOGRAM;

 ods listing style=journal;
 ods listing  gpath='d:/png';
 ods graphics on / width=4in
  imagefmt=png imagename="sbys";

 proc sgplot data=sashelp.class;
 format age 2.;
 vbar age /datalabel;
 run;quit;

 ods graphics off;

 * ADD HISTOGRAM RIGHT OF REPORT;

 %utl_submit_r64('
 library(openxlsx);
 wb<-loadWorkbook(
     file="d:/xls/sbys.xlsx");
 insertImage(
   wb
   ,"sheet1"
   ,"d:/png/sbys.png"
   ,width    = 8
   ,height   = 6
   ,startRow = 1
   ,startCol = "E"
   ,units = "cm");
 saveWorkbook(
     wb
   ,"d:/xls/sbys.xlsx"
   ,overwrite = TRUE);
 ');

/*           _               _
  ___  _   _| |_ _ __  _   _| |_
 / _ \| | | | __| `_ \| | | | __|
| (_) | |_| | |_| |_) | |_| | |_
 \___/ \__,_|\__| .__/ \__,_|\__|
                |_|
*/

/**************************************************************************************************************************/
/*                                                                                                                        */
/*                                                                                                                        */
/*                   OUTPUT(d:/x;s/sbys.xlsx)                                                                             */
/*                                                                                                                        */
/*    EXCEL   A       B        C         PRODUCTS                                                                         */
/*     ROW -------|--------|--------                                                                                      */
/*                                     ---+--+--+--+--                                                                    */
/*     1   Country  Product  Actual   6+             +6                                                                   */
/*                                     |             |                                                                    */
/*     2   CANADA   BED      $47,729  4+       ** ** +4                                                                   */
/*     3            CHAIR    $50,239   |       ** ** |                                                                    */
/*     4            DESK     $52,187  2+ ** ** ** ** +2                                                                   */
/*     5            SOFA     $50,135   | ** ** ** ** |                                                                    */
/*     6            TABLE    $46,700   ---+--+--+--+--                                                                    */
/*     7   GERMANY  BED      $46,134      B  C  D  T                                                                      */
/*     8           CHAIR     $47,105      E  H  E  A                                                                      */
/*     9            DESK     $48,502      D  A  S  B                                                                      */
/*                                           I  K  L                                                                      */
/*     ------                                R                                                                            */
/*     SHEET1                                                                                                             */
/*     ------                                                                                                             */
/*                                                                                                                        */
/**************************************************************************************************************************/

EPO
-----------------------------------------------------------------------------------------------------------------------------------
ttps://github.com/rogerjdeangelis/utl-Create-a-side-by-side-table-and-graph-using-greplay
ttps://github.com/rogerjdeangelis/utl-Graph-with-known-intercept-and-slope
ttps://github.com/rogerjdeangelis/utl-Graphics-Surveying-ten-random-locations-in-North-Carolina-using-superinposed-grid
ttps://github.com/rogerjdeangelis/utl-NHANES-full-raw-demographic-and-health-data-R-Package
ttps://github.com/rogerjdeangelis/utl-R-AI-igraph-list-connections-in-a-non-directed-graph-for-a-subset-of-vertices
ttps://github.com/rogerjdeangelis/utl-adding-text-to-an-existing-png-graphic-python-AI
ttps://github.com/rogerjdeangelis/utl-changepoint-like-analysis-in-R-and-SAS-elbow-graph
ttps://github.com/rogerjdeangelis/utl-classic-sas-and-well-designed-tables-and-ascii-graphics-instead-of-bling
ttps://github.com/rogerjdeangelis/utl-color-a-region-under-a-distribution-curve-graph-wps-and-wps-r
ttps://github.com/rogerjdeangelis/utl-create-graphs-in-excel-using-excel-chart-templates
ttps://github.com/rogerjdeangelis/utl-creating-a-clinical-demographic-report-using-r-and-python-sql
ttps://github.com/rogerjdeangelis/utl-dygraph-javascript-library-from-MIT
ttps://github.com/rogerjdeangelis/utl-excel-report-with-two-side-by-side-graphs-below_python
ttps://github.com/rogerjdeangelis/utl-graphics-boxplots-with-jiggled-point-values-alongside
ttps://github.com/rogerjdeangelis/utl-how-many-triangles-in-the-polygon-r-igraph-AI
ttps://github.com/rogerjdeangelis/utl-identical-side-by-side-text-and-graphics-in-pdf-and-powerpoint
ttps://github.com/rogerjdeangelis/utl-identify-linked-and-unliked-paths-r-igraph
ttps://github.com/rogerjdeangelis/utl-igraph-find-largest-group-of-unrelated-individuals-in-your-family-reunion
ttps://github.com/rogerjdeangelis/utl-make-fake-relational-clinical-tables-demographics-lab-exposure-adverse-events
ttps://github.com/rogerjdeangelis/utl-quality-graphics-in-R-wps-and-sas
ttps://github.com/rogerjdeangelis/utl-r-graphics-vs-wps-base-graphics-layering-in-r-versus-procs-in-wps-base-ggplot2
ttps://github.com/rogerjdeangelis/utl-shortest-and-longest-travel-time-from-home-to-work-igraph-AI
ttps://github.com/rogerjdeangelis/utl-three-heat-maps-of-bivariate-normal-wps-r-graph-plot
ttps://github.com/rogerjdeangelis/utl-under-used-proc-calendar-ascii-graphics
ttps://github.com/rogerjdeangelis/utl_R_graphics_polar_plot
ttps://github.com/rogerjdeangelis/utl_adding_SAS_graphics_at_an_arbitrary_position_into_existing_excel_sheets
ttps://github.com/rogerjdeangelis/utl_classic_sas_graph_greplay_harvard_macro_multiple_plots_per_page
ttps://github.com/rogerjdeangelis/utl_classic_sas_graph_three_plots_across_many_methods_long_live
ttps://github.com/rogerjdeangelis/utl_custom_graphics_in_R
ttps://github.com/rogerjdeangelis/utl_download_2015_ACS_5yr_zipcode_level_american_community_survey_demographics_as_sas_dataset
ttps://github.com/rogerjdeangelis/utl_fun_with_line_printer_graphics_editor
ttps://github.com/rogerjdeangelis/utl_graph_visualize_crosstab
ttps://github.com/rogerjdeangelis/utl_graphics_589_SAS_and_R_graphics_with_code_and_datasets
ttps://github.com/rogerjdeangelis/utl_graphics_determine_us_state_from_latitude_and_longitude
ttps://github.com/rogerjdeangelis/utl_graphics_fit_a_smooth_line_to_a_scatter_plot_loess
ttps://github.com/rogerjdeangelis/utl_graphics_flexibility_of_ascii_bar_charts
ttps://github.com/rogerjdeangelis/utl_graphics_plotting_rivers_in_brazil_using_sharpefiles
ttps://github.com/rogerjdeangelis/utl_graphics_zipcode_boundary_maps
ttps://github.com/rogerjdeangelis/utl_how_to_stack_a_table_and_corresponding_bar_graph
ttps://github.com/rogerjdeangelis/utl_javascript_and_classic_map_graphics_with_mouseovers_and_multiple_drilldowns
ttps://github.com/rogerjdeangelis/utl_javascript_dygraph_graphics_multipanel_time_series
ttps://github.com/rogerjdeangelis/utl_minimal_code_for_demographic_clinical_n_percent_report
ttps://github.com/rogerjdeangelis/utl_pdf_graphics_top_40_a_sas_ods_graphics_look_at_chicago_public_schools_salaries_by_job
ttps://github.com/rogerjdeangelis/utl_polar_graphics_pot_violin_plot
ttps://github.com/rogerjdeangelis/utl_proc_gmap_classic_graphics_grid_containing_four_states
ttps://github.com/rogerjdeangelis/utl_r_graphics_visualizing_assciation_amoung_many_variables
ttps://github.com/rogerjdeangelis/utl_remove_isolated_nodes_from_an_network_r_igraph
ttps://github.com/rogerjdeangelis/utl_sas_classic_graphics_15_plots_on_a_page
ttps://github.com/rogerjdeangelis/utl_sas_classic_graphics_designing_your_greplay_template
ttps://github.com/rogerjdeangelis/utl_sas_classic_graphics_grid_of__proc_univariate_histograms
ttps://github.com/rogerjdeangelis/utl_sas_classic_graphs_using_phil_mason_grid_macro_for_layout
ttps://github.com/rogerjdeangelis/utl_table_graph_ppt
ttps://github.com/rogerjdeangelis/utl_wps_sas_classic_graphics_optimum_minimums_maximums_increments_for-axes


/*              _
  ___ _ __   __| |
 / _ \ `_ \ / _` |
|  __/ | | | (_| |
 \___|_| |_|\__,_|

*/
