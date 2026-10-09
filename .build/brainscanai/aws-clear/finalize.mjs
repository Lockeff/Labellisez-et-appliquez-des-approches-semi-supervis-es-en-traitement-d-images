import fs from 'node:fs/promises';
import path from 'node:path';
import {createHash} from 'node:crypto';
import {pathToFileURL} from 'node:url';
const workspaceDir=process.cwd();
const skill='C:/Users/ronan/.codex/plugins/cache/openai-primary-runtime/presentations/26.904.11930/skills/presentations';
const validators=path.join(skill,'container_tools');
const {finalizePresentation}=await import(pathToFileURL(path.join(validators,'artifact_tool_utils.mjs')).href);
const referencePath=path.join(workspaceDir,'.build/brainscanai/aws-clear/source-live.pptx');
const referenceSha256=createHash('sha256').update(await fs.readFile(referencePath)).digest('hex');
const result=await finalizePresentation({
  workspaceDir,
  candidatePath:path.join(workspaceDir,'.build/brainscanai/aws-clear/candidate-v4.pptx'),
  finalPath:path.join(workspaceDir,'output/presentations/BrainScanAI_presentation_13_slides_couts.pptx'),
  pythonExecutable:path.join(workspaceDir,'.venv/Scripts/python.exe'),
  integrityValidatorPath:path.join(validators,'inspect_presentation_package_integrity.py'),
  layoutValidatorPath:path.join(validators,'inspect_presentation_layout_geometry.py'),
  layoutArgs:['--expected-slide-size-emu','12192000,6858000','--validate-heading-fit','--validate-bullet-geometry','--require-native-table-slide','11','--require-native-table-slide','12','--require-native-table-slide','13'],
  explicitTotalSlideCount:13,
  requiredNativeTableOwnerSlides:[11,12,13],
  requiredNativeChartOwnerSlides:[5],
  nativeChartTargetApplication:'portable',
  fontPolicy:{basis:'reference',families:['Arial','Aptos'],referencePath,referenceSha256},
  verifyArtifactToolImport:false,
  receiptPath:path.join(workspaceDir,'.build/brainscanai/aws-clear/final-validation-v2.json')
});
console.log(JSON.stringify({finalPath:result.finalPath,slides:result.packageIntegrity.slide_count,layoutFindings:result.presentationLayout.findingCount,integrity:result.packageIntegrity.status,chartPassed:result.nativeChartValidation.passed},null,2));
