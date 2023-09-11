import { OBC, THREE } from "./base.js";

export const initialize = () => {
	const container = document.getElementById("container");
	if (!container) return;
	const components = new OBC.Components();

	components.scene = new OBC.SimpleScene(components);
	components.renderer = new OBC.PostproductionRenderer(components, container);
	components.camera = new OBC.SimpleCamera(components);
	components.raycaster = new OBC.SimpleRaycaster(components);

	components.init();

	components.renderer.postproduction.enabled = true;

	const scene = components.scene.get();

	components.camera.controls.setLookAt(10, 10, 10, 0, 0, 0);

	const directionalLight = new THREE.DirectionalLight();
	directionalLight.position.set(5, 10, 3);
	directionalLight.intensity = 0.5;
	scene.add(directionalLight);

	const ambientLight = new THREE.AmbientLight();
	ambientLight.intensity = 0.5;
	scene.add(ambientLight);

	const grid = new OBC.SimpleGrid(components, new THREE.Color(0x666666));
	components.tools.add("grid", grid);
	const gridMesh = grid.get();
	const effects = components.renderer.postproduction.customEffects;
	effects.excludedMeshes.push(gridMesh);

	// now every is OK, let see in add in
	// addin is OK
	// now add a button in a bottom panel

	const fragments = new OBC.FragmentManager(components);

	const highlighter = new OBC.FragmentHighlighter(components, fragments);
	// load model and update highlighter

	const toolbar = new OBC.Toolbar(components);
	components.ui.addToolbar(toolbar);

	const loadButton = new OBC.Button(components);
	loadButton.materialIcon = "download";
	loadButton.tooltip = "Load model";
	toolbar.addChild(loadButton);
	loadButton.onclick = () => loadModel();

	function loadModel() {
		loadFragment(components, fragments, highlighter);
	}

	components.renderer.postproduction.customEffects.outlineEnabled = true;
	highlighter.outlinesEnabled = true;

	const highlightMaterial = new THREE.MeshBasicMaterial({
		color: "#BCF124",
		depthTest: false,
		opacity: 0.8,
		transparent: true,
	});

	highlighter.add("default", highlightMaterial);
	highlighter.outlineMaterial.color.set(0xf0ff7a);

	let lastSelection;

	let singleSelection = {
		value: true,
	};

	function highlightOnClick(event) {
		const result = highlighter.highlight("default", singleSelection.value);
		if (result) {
			highlightHighestValue(result.id);
			lastSelection = {};
			for (const fragment of result.fragments) {
				const fragmentID = fragment.id;
				lastSelection[fragmentID] = [result.id];
			}
		}
	}

	container.addEventListener("click", (event) => highlightOnClick(event));
};

// load fragment file first
function loadFragment(components, fragments, highlighter) {
	(async () => {
		// we missing method and cors
		const resFrag = await fetch("http://localhost:3000/download/CenterConference.frag", {
			method: "POST",
			mode: "cors",
		});
		if (resFrag.ok) {
			resFrag.arrayBuffer().then((dataBlob) => {
				const buffer = new Uint8Array(dataBlob);
				const model = fragments.load(buffer);
				highlighter.update();
				// OK now fit to zoom model
				fitToZoom(components, model);
				loadData();
			});
		} else {
			const res = resFrag.json();
			console.log(res);
		}
	})();
}

// fit to m zoom model
function fitToZoom(components, model) {
	const { max, min } = model.boundingBox;
	if (!max || !min) return;
	// define vector from max to min
	const dir = max.clone().sub(min.clone()).normalize();
	// distance max to min
	const dis = max.distanceTo(min);
	// center

	const center = max.clone().add(dir.clone().multiplyScalar(-0.5 * dis));

	// camera position
	const pos = max.clone().add(dir.clone().multiplyScalar(0.5 * dis));
	// set true mean we can animate
	components.camera.controls.setLookAt(pos.x, pos.y, pos.z, center.x, center.y, center.z, true);
}
// test in addin
// load data
// ok, but we need a button to load model because the app will clash
// means we need an event for that
// we have to shut dow the test in broseer because this.....
// but when we deploy our app, will not

function loadData() {
	(async () => {
		// we missing method and cors
		const resFrag = await fetch("http://localhost:3000/download/CenterConference.json", {
			method: "POST",
			mode: "cors",
		});
		if (resFrag.ok) {
			resFrag.json().then((jsonData) => {
				loadDataToExcel(computeData(jsonData));
			});
		} else {
			const res = resFrag.json();
			console.log(res);
		}
	})();
}
// when we export json, i had added more data here
//spatialTree
// buildingStorey 1
// children .....
// buildingStorey 2
// children .....
// buildingStorey 3
// children .....
// buildingStorey 4
// children .....
// buildingStorey 5
// children .....

// siqualize data to put to table on excel

// because childID:buildingStoreyID
function computeData(jsonData) {
	// first we have to storey all json by level
	if (!jsonData.spatialTree) return null;
	const buildings = {};
	const spatialTree = jsonData.spatialTree;
	const title = {};
	const elements = [];
	Object.keys(spatialTree).forEach((key) => {
		if (!buildings[spatialTree[key]]) {
			buildings[spatialTree[key]] = {
				buildingStorey: jsonData[spatialTree[key]],
				children: [],
			};
		}
		// we need add more params : Elevation , buildingStorey name, buildingStorey ID
		const element = jsonData[key];
		const building = buildings[spatialTree[key]].buildingStorey;
		element.Elevation = building.Elevation.value;
		// name, if find ObjectTpye, then Name, then LongName
		element.buildingName = building.ObjectType
			? building.ObjectType.value
			: building.Name
			? building.Name.value
			: building.LongName
			? building.LongName.value
			: "";
		element.buildingStoreyId = building.expressID;
		Object.keys(element).forEach((key) => {
			if (!title[key]) title[key] = key;
		});
		elements.push(element);
		// now we have to sotorage all params
		// we need all object
	});
	const values = [];
	// we need a header
	const headers = Object.keys(title);
	values.push(headers);
	elements.forEach((element) => {
		values.push(
			headers.map((h) => {
				return element[h]?.value || element[h] || "";
			})
		);
	});
	return values;
}
// oK now we have to add to table

function loadDataToExcel(values) {
	if (!Excel) return;
	Excel.run(function (ctx) {
		// Create a proxy object for the active sheet
		var sheet = ctx.workbook.worksheets.getActiveWorksheet();
		sheet.name = "Elements";
		// Assuming 'values' is a two-dimensional array (rows and columns)
		var numRows = values.length;
		var numCols = values[0].length;
		// Calculate the target range dynamically based on the size of 'values'
		var targetRange = sheet.getRangeByIndexes(0, 0, numRows, numCols); // Start at A1 (0,0)
		// Queue a command to write the sample data to the worksheet
		targetRange.values = values;
		// Freeze the first row
		sheet.freezePanes.freezeRows(1);

		targetRange.getRow(0).format.font.bold = true;
		targetRange.getRow(0).format.fill.color = "green";
		// Apply a border around the entire range
		targetRange.format.borders.getItem("EdgeTop").style = "Continuous";
		targetRange.format.borders.getItem("EdgeTop").color = "Black";
		targetRange.format.borders.getItem("EdgeBottom").style = "Continuous";
		targetRange.format.borders.getItem("EdgeBottom").color = "Black";
		targetRange.format.borders.getItem("EdgeLeft").style = "Continuous";
		targetRange.format.borders.getItem("EdgeLeft").color = "Black";
		targetRange.format.borders.getItem("EdgeRight").style = "Continuous";
		targetRange.format.borders.getItem("EdgeRight").color = "Black";
		// Run the queued-up commands, and return a promise to indicate task completion
		targetRange.getEntireColumn().format.autofitColumns();
		return ctx.sync();
	}).catch(errorHandler);
}
// OK now is highlight

// idea we find A1 same expessID
// now we have expressID every we select element
function highlightHighestValue(expressID) {
	// Run a batch operation against the Excel object model
	if (!Excel) return;
	Excel.run(function (ctx) {
		// Create a proxy object for the used range of the worksheet and load its properties
		var usedRange = ctx.workbook.worksheets.getActiveWorksheet().getUsedRange().load("values");

		// Run the queued-up command, and return a promise to indicate task completion
		return ctx
			.sync()
			.then(function () {
				var cellToHighlight = null;

				// Use the 'find' method to search for 'expressID' in column A (assuming it's in the first column, A)
				cellToHighlight = usedRange.find(expressID, { matchWholeCell: true, matchCase: false });

				if (cellToHighlight) {
					// Clear any previous highlighting
					usedRange.worksheet.getUsedRange().format.fill.clear();
					usedRange.worksheet.getUsedRange().format.font.bold = false;

					// Highlight the matching row
					var rowToHighlight = cellToHighlight.getEntireRow();
					//rowToHighlight.format.fill.color = "orange";
					//rowToHighlight.format.font.bold = true;
					rowToHighlight.select();
				}
			})
			.then(ctx.sync);
	}).catch(errorHandler);
}

// expressID  type      level ....
// 1					IFCWALL
// 2					IFCWALL
// 3					IFCWALL
//.......

// all finish, however we can integrate more like when we select on table , the viewer also select on element
// that's all
// enjoy
