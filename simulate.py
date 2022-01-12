import com3dx
CATIA = get3dxClient()
ENUM = win32com.client.constants

myEditor = CATIA.ActiveEditor
myProdService = myEditor.GetService("PLMProductService")
myEntities = myProdService.EditedContent
myEntity = myEntities.Item(1)
MySimulationRoot = myEntity
myModel = MySimulationRoot.Model

myPublications = myModel.Publications
myPublication = myPublications.GetItem("Part_Body_Publication_Name")
myPrdRepFactory = myModel.GetItem("SimPrdRepFactory")
	myRepRef = myPrdRepFactory.CreatePrdRep("FEM")
myFEMRoot = myRepRef.GetItem("SimFemRoot")
	myFEMRoot.AddAssociatedRep(myPublication)
    myMeshSet = myFEMRoot.GetSet("SimNodesElements")
	myMeshParts = myMeshSet.MeshParts
	myMeshPart = myMeshParts.Add("CATFmtOctree3DRulesMesher")
	myMeshPart.AddSupport(myPublication)
	myMeshPart.Update
