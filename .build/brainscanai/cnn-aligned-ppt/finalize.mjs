import fs from 'node:fs/promises';
import path from 'node:path';
import {createHash} from 'node:crypto';
import {pathToFileURL} from 'node:url';
const workspaceDir=process.cwd();
const build=path.join(workspaceDir,'.build/brainscanai/cnn-aligned-ppt');
const validators='C:/Users/ronan/.codex/plugins/cache/openai-primary-runtime/presentations/26.904.11930/skills/presentations/container_tools';
const {finalizePresentation}=await import(pathToFileURL(path.join(validators,'artifact_tool_utils.mjs')).href);
const source=path.join(workspaceDir,'output/presentations/BrainScanAI_presentation_V1_Final.pptx');
const candidate=path.join(build,'candidate.pptx');
const checks=JSON.parse(await fs.readFile(path.join(build,'scope-math-validation.json'),'utf8'));
const sha=p=>fs.readFile(p).then(b=>createHash('sha256').update(b).digest('hex'));
if(!checks.passed || checks.candidate_sha256!==await sha(candidate)) throw new Error('Scope/math checks required on current candidate');
const tables=[2,3,5,6,7,9];
const result=await finalizePresentation({
 workspaceDir,
 candidatePath:candidate,
 finalPath:path.join(workspaceDir,'output/presentations/BrainScanAI_presentation_V1_Final_actualisee.pptx'),
 pythonExecutable:path.join(workspaceDir,'.venv/Scripts/python.exe'),
 integrityValidatorPath:path.join(validators,'inspect_presentation_package_integrity.py'),
 layoutValidatorPath:path.join(validators,'inspect_presentation_layout_geometry.py'),
 layoutArgs:['--expected-slide-size-emu','12192000,6858000','--validate-heading-fit','--validate-bullet-geometry',...tables.flatMap(i=>['--require-native-table-slide',String(i)])],
 explicitTotalSlideCount:13,
 requiredNativeTableOwnerSlides:tables,
 requiredNativeChartOwnerSlides:[5],
 nativeChartTargetApplication:'portable',
 fontPolicy:{basis:'reference',families:['Arial','Aptos'],referencePath:source,referenceSha256:await sha(source)},
 verifyArtifactToolImport:false,
 receiptPath:path.join(build,'final-validation.json')
});
if(await sha(candidate)!==await sha(result.finalPath)) throw new Error('Final export differs from rendered and reviewed candidate');
console.log(JSON.stringify({finalPath:result.finalPath,slideCount:result.packageIntegrity.slide_count,integrity:result.packageIntegrity.status,layoutFindings:result.presentationLayout.findingCount},null,2));
