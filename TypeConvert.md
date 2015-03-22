# Introduction #

TypeConvert is the ArcGIS extension allows you to convert one type feature class to another type feature class. Also we recommend to use module <a href='http://kmler.geoblogspot.com/'><b>KMLer</b></a> as an extension for professional work with Google Earth.<br>
<br>
<img src='https://sites.google.com/a/geoblogspot.com/typeconvert/Home/tctoolsall-full.jpg'><br>
<br>
Flash presentations:<br>
<a href='http://www.dataplus.ru/Soft/PGI/typeconvertgraph.htm'>How to use</a><br>
<a href='http://www.dataplus.ru/Soft/PGI/mytoolboxinstall.htm'>Install geoprocessing tools</a><br>

<br>
<b>Functionality</b><br>
<ul><li><b>Polygon-Polyline-Points</b><br>
Convert one type feature class to another type feature class<br>
<li><b>To ConvexHull</b> and <b>To envelope</b><br>
Convert feature class into convex hull or envelope. In this case new polygon feature<br>
class containing only one feature (convex hull or envelope) will be created<br>
<li><b>To centroid</b><br>
Command allows you to create a new point feature class with all attributes from the center points (centroids) of the features in the current layer. A centroid of a feature is the spatial location of its envelope center. Additional fields 'Z', 'Zmin' and 'Zmax' are created if feature class with MZ geometries is converted.<br>
<li><b>To segments</b><br>
Command allow you to convert polyline or polygon feature class to polyline feature class, consisting of lines segments of an initial feature class.<br>
<li><b>From graphics</b><br>
Command allows you to convert graphics elements of the active map into features. New created features are stored in corresponding feature classes, according to their geometry types. Names of new classes consists of a name, set by the user and a suffix, indicating geometry type (e.g. "classname_polyline"). Text associated with the graphics is stored into 'Text' field of the target feature class.<br>
<li><b>Remove duplicates</b><br>
Command allows you to remove duplicate features from the current layer. All 'cleaned' features are stored in output feature class. All 'removed' features are stored in<br>
additional output feature class with "<i>duplicate" suffix. Two features are compared by coordinates only without taking into account any difference in attributes.<br></i><li><b>Stratification</b><br>
<br>
<img src='http://xbbster.googlepages.com/stratshema1.gif' alt='http://xbbster.googlepages.com/stratshema1.gif'><br>
<br>
Command allows you to stratify current layer, classified by categories or quantities, to set of layers according the current legend. The layer must has symbology, based on single field. Names, aliases of the new feature classes are based on a attribute name, value of a class and a label of a class. All created feature classes keep attributes of source feature class. Styles of classes and categories are kept in new layers.<br>
<li><b>Divide segments</b><br>
Add points to long segments of current feature class.<br>
<li><b>To <b>.bln</b></b><br>
Command allows you to export feature class with 2-dimentional geometry (polyline, polygon) to blanking file (<b>.bln) for using this file in Golden Software Surfer.<br></b><li><b>To Google Earth</b><br>
Command exports feature class to KML-file for using this file in Google Earth. If you have <a href='http://earth.google.com/'>Google Earth</a> installed, all you need is just click on the <b>.kml file you have created for a superb visualization of the results in 3D.</b>

<li><b>About</b><br>
Opens the window of the information on the current version of the program.<br>
<br>
<b>Notes</b>
<ul><li>You can project new feature class to data frame coordinate system.<br>
<li>You can convert selected objects (features) only.<br>
<li>You can convert all objects (features) in view extent.<br>
<li>You can convert feature class into the same feature class.<br>
<li>If a point feature class is converted, a new two special fields("gsGroup","gsOrder") will be created. 'gsGroup' field contains a group number. <br>
'gsOrder' field contains an order number in group. <br>
It's very important for polyline feature class creatation. You should specify group number and object number in group for each object(feature) in point feature class.<br>
<li>TypeConvert works with current selected layer. If more than one layer is selected, it works with the first selected layer.<br>
<li>TypeConvert works with multipart objects and MZ geometries.<br>
<br><br><br>
If you have problems by installation then see <a href='http://applications.geoblogspot.com/deployment--requirements'>this page</a>.<br>
To install the geoprocessing tools:<br>
1. Open ArcCatalog and search folder of Typeconvert. By default it is<br>
C:\Program Files\GISCenter\Typeconvert\<br>
2. You see toolbox Typeconvert. Right click on toolbox and select "Add to Toolbox". The Typeconvert toolbox stored in Toolbox.