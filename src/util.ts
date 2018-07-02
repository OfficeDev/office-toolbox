/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

import * as ai from 'applicationinsights';
import * as chalk from 'chalk';
import * as fs from 'fs-extra';
import * as jszip from 'jszip';
import * as junk from 'junk';
import * as officeAddinValidator from 'office-addin-validator';
import * as opn from 'opn';
import * as os from 'os';
import * as path from 'path';
import * as shell from 'node-powershell';
import * as xml2js from 'xml2js';

export const appInsights = ai.getClient('7695b3c1-32c5-4458-99d6-5d0e3208c9c2');

const office16RegistryPath = 'HKCU:\\Software\\Microsoft\\Office\\16.0';
const wefFolder = '\\WEF';
const developerFolder = '\\Developer';

export const applicationProperties = {
  word: {
    TaskPaneApp: {
      webExtensionPath: 'word/webextensions/webextension.xml',
      templateName: 'DocumentWithTaskPane.docx'
    },
    sideloadingDirectory: path.join(os.homedir(), 'Library/Containers/com.microsoft.Word/Data/Documents/wef'),
    documentationLink: "https://docs.microsoft.com/en-us/office/dev/add-ins/testing/create-a-network-shared-folder-catalog-for-task-pane-and-content-add-ins",
    canSideload: true
  },
  excel: {
    TaskPaneApp: {
      webExtensionPath: 'xl/webextensions/webextension.xml',
      templateName: 'BookWithTaskPane.xlsx'
    },
    ContentApp: {
      webExtensionPath: 'xl/webextensions/webextension.xml',
      templateName: 'BookWithContent.xlsx'
    },
    sideloadingDirectory: path.join(os.homedir(), 'Library/Containers/com.microsoft.Excel/Data/Documents/wef'),
    documentationLink: "https://docs.microsoft.com/en-us/office/dev/add-ins/testing/create-a-network-shared-folder-catalog-for-task-pane-and-content-add-ins",
    canSideload: true
  },
  powerpoint: {
    TaskPaneApp: {
      webExtensionPath: 'ppt/webextensions/webextension.xml',
      templateName: 'PresentationWithTaskPane.pptx'
    },
    ContentApp: {
      webExtensionPath: 'ppt/slides/udata/data.xml',
      templateName: 'PresentationWithContent.pptx'
    },
    sideloadingDirectory: path.join(os.homedir(), 'Library/Containers/com.microsoft.Powerpoint/Data/Documents/wef'),
    documentationLink: "https://docs.microsoft.com/en-us/office/dev/add-ins/testing/create-a-network-shared-folder-catalog-for-task-pane-and-content-add-ins",
    canSideload: true
  },
  outlook: {
    documentationLink: "https://docs.microsoft.com/en-us/outlook/add-ins/sideload-outlook-add-ins-for-testing",
    canSideload: false
  },
  onenote: {
    documentationLink: "https://docs.microsoft.com/en-us/office/dev/add-ins/onenote/onenote-add-ins-getting-started",
    canSideload: false
  },
  project: {
    documentationLink: "https://docs.microsoft.com/en-us/office/dev/add-ins/project/project-add-ins",
    canSideload: false
  }
};

// TOP-LEVEL COMMANDS //
export function sideload(application: string, manifestPath: string): Promise<any> {
  appInsights.trackEvent('sideload');
  return sideloadManifest(application, manifestPath);
}

export function list(): Promise<Array<[string, string, string]>> {
  appInsights.trackEvent('list');
  return getAllIdsAndManifests();
}

export function remove(application: string, manifestPath: string): Promise<any> {
  appInsights.trackEvent('remove');
  return removeManifest(application, manifestPath);
}

export function validate(manifestPath: string): Promise<string> {
  appInsights.trackEvent('validate');
  return validateManifest(manifestPath);
}

// DISAMBIGUATING COMMANDS //
export function getManifests(application: string): Promise<string[]> {
  return (process.platform === 'win32')
    ? getManifestsFromRegistry()
    : getManifestsFromSideloadingDirectory(application);
}

function addManifest(application: string, manifestPath: string): Promise<any> {
  return (process.platform === 'win32')
    ? addManifestToRegistry(manifestPath)
    : addManifestToSideloadingDirectory(application, manifestPath);
}

async function getAllManifests(): Promise<string[]> {
  if (process.platform === 'win32') {
    return getManifestsFromRegistry();
  }
  else {
    let manifests = [];
    for (const application of Object.keys(applicationProperties)) {
      manifests = [...manifests, await getManifests(application)];
    }
    return Promise.resolve(manifests);
  }
}

function removeManifest(application: string, manifestPath: string): Promise<any> {
  if (fs.existsSync(manifestPath)) {
    manifestPath = fs.realpathSync(manifestPath);
  }

  return (process.platform === 'win32'
    ? removeManifestFromRegistry(manifestPath)
    : removeManifestFromSideloadingDirectory(application, manifestPath));
}

// NON-WIN32 COMMANDS //
function addManifestToSideloadingDirectory(application: string, manifestPath: string): Promise<any> {
  return new Promise ((resolve, reject) => {
    const sideloadingDirectory = applicationProperties[application].sideloadingDirectory;
    fs.ensureDirSync(sideloadingDirectory);

    const sideloadingManifestPath = path.join(sideloadingDirectory, path.basename(manifestPath));

    if (fs.existsSync(sideloadingManifestPath)) {
      const stat = fs.statSync(manifestPath);
      const sideloadingStat = fs.statSync(sideloadingManifestPath);

      if (stat.ino !== sideloadingStat.ino && stat.dev !== sideloadingStat.dev) {
        return reject(['Remove the manifest with matching name before adding this one. ', fs.realpathSync(sideloadingManifestPath)]);
      }
    }

    fs.ensureLinkSync(manifestPath, sideloadingManifestPath);
    resolve();
  });
}

function getManifestsFromSideloadingDirectory(inputApplication: string): Promise<string[]> {
  return new Promise((resolve, reject) => {
    let manifestPaths = [];

    for (let application of Object.keys(applicationProperties)) {
      if (!inputApplication || application === inputApplication) {
        const sideloadingDirectory = applicationProperties[application].sideloadingDirectory;

        if (!fs.existsSync(sideloadingDirectory)) {
          continue;
        }

        fs.readdirSync(sideloadingDirectory).filter(junk.not).forEach(manifestName => {
          manifestPaths.push(fs.realpathSync(path.join(sideloadingDirectory, manifestName)));
        });
      }
    }
    resolve(manifestPaths);
  });
}

function removeManifestFromSideloadingDirectory(inputApplication: string, manifestPathToRemove: string): Promise<any> {
  return new Promise(async (resolve, reject) => {
    try {
      let manifestRemoved = false;

      for (let application of Object.keys(applicationProperties)) {
        if (!inputApplication || application === inputApplication) {
          const sideloadingDirectory = applicationProperties[application].sideloadingDirectory;

          if (!fs.existsSync(sideloadingDirectory)) {
            continue;
          }

          fs.readdirSync(sideloadingDirectory).forEach(manifestName => {
            const realManifestPath = (fs.realpathSync(path.join(sideloadingDirectory, manifestName)));
            if (manifestPathToRemove === realManifestPath) {
              console.log(`Removing ${manifestPathToRemove} for application ${application}`);
              fs.unlinkSync(manifestPathToRemove);
              manifestRemoved = true;
            }
          });
        }
      }

      if (!manifestRemoved) {
        return reject('No manifests were found to remove. Use "list" to show manifests that have been added.');
      }
    } catch (err) {
      reject(err);
    }
  });
}

// WIN32 SPECIFIC COMMANDS //
function querySideloadingRegistry(commands: string[]): Promise<string> {
  return new Promise(async (resolve, reject) => {
    let ps = new shell({'debugMsg': false});

    try {
      // Ensure that the registry path exists
      ps.addCommand('$RegistryPath = "' + office16RegistryPath + '"');
      ps.addCommand('if(!(Test-Path $RegistryPath)) { Throw "NO-OFFICE-16" } ');
      ps.addCommand('$RegistryPath = "' + office16RegistryPath + wefFolder + '"');
      ps.addCommand('if(!(Test-Path $RegistryPath)) { New-Item -Path $RegistryPath } ');
      ps.addCommand('$RegistryPath = "' + office16RegistryPath + wefFolder + developerFolder + '"');
      ps.addCommand('if(!(Test-Path $RegistryPath)) { New-Item -Path $RegistryPath } ');

      await ps.invoke();
      ps.dispose();

      ps = new shell({'debugMsg': false});
      ps.addCommand('$RegistryPath = "' + office16RegistryPath + wefFolder + developerFolder + '"');
      for (const command of commands) {
        ps.addCommand(command);
      }

      const output = await ps.invoke();
      ps.dispose();
      resolve(output);
    } catch (err) {
      ps.dispose();
      if (err.indexOf('NO-OFFICE-16') > -1) {
        reject(office16RegistryPath + ' could not be found in the registry. Make sure Microsoft Office is installed.');
      }
      else {
        reject(err);
      }
    }
  });
}

function addManifestToRegistry(manifestPath: string): Promise<any> {
  return querySideloadingRegistry(['Set-ItemProperty -LiteralPath $RegistryPath -Name "' + manifestPath + '" -Value "' + manifestPath + '"']);
}

function getManifestsFromRegistry(): Promise<string[]> {
  return new Promise(async(resolve, reject) => {
    try {
      const registryOutput = await querySideloadingRegistry(['Get-ItemProperty -LiteralPath $RegistryPath | ConvertTo-Json -Compress']);

      if (!registryOutput || registryOutput.indexOf('{') === -1) {
        resolve([]);
      }

      const registryJSON = JSON.parse(registryOutput);
      let manifestPaths = [];

      for (const name in registryJSON) {
        // Manifests are inserted in the registry with matching name and value
        if (registryJSON[name].toString().toLowerCase() === name.toString().toLowerCase()) {
          manifestPaths.push(name);
        }
      }

      resolve(manifestPaths);
    } catch (err) {
      reject(err);
    }
  });
}

function removeManifestFromRegistry(manifestPath: string): Promise<any> {
  if (!manifestPath) {
    return new Promise((resolve, reject) => {
      return reject('No manifest was specified');
    });
  }
  console.log(`Removing ${manifestPath}`);
  return querySideloadingRegistry(['Remove-ItemProperty -LiteralPath $RegistryPath -Name "' + manifestPath + '" -ErrorAction SilentlyContinue']);
}

// GENERIC HELPER FUNCTIONS //
function isGuid (text: string): boolean {
  const guidRegex = /^[0-9A-F]{8}[-]?([0-9A-F]{4}[-]?){3}[0-9A-F]{12}?/i;
  return guidRegex.test(text);
}

function sideloadManifest(application: string, manifestPath: string): Promise<any> {
  return new Promise(async (resolve, reject) => {
    try {
      if (fs.existsSync(manifestPath) && fs.lstatSync(manifestPath).isFile()) {
        manifestPath = fs.realpathSync(manifestPath);
      }
      else {
        return reject(['The manifest to sideload could not be found: ', manifestPath]);
      }

      const [parsedType, parsedGuid, parsedVersion] = await parseManifest(manifestPath);
      await addManifest(application, manifestPath);
      const templateFile = await generateTemplateFile(application, parsedType, parsedGuid, parsedVersion);

      appInsights.trackEvent('open', { guid: parsedGuid, version: parsedVersion });
      console.log(`Opening file ${templateFile}`);
      opn(templateFile);
      resolve();
    }
    catch (err) {
      return reject(err);
    }
  });
}

function getAllIdsAndManifests(): Promise<Array<[string, string, string]>> {
  let applications = [];
  if (process.platform === 'win32') {
    applications = [null];
  }
  else {
    applications = Object.keys(applicationProperties);
  }

  return new Promise(async (resolve, reject) => {
    try {
      let allIdsAndManifests = [];

      for (const application of applications) {
        const idsAndManifests = await getIdsAndManifests(application);
        for (const [id, manifest] of idsAndManifests) {
          allIdsAndManifests.push([id, manifest, application]);
        }
      }

      resolve(allIdsAndManifests);
    } catch (err) {
      return reject(err);
    }
  });
}

function validateManifest(manifestPath: string): Promise<string> {
  return new Promise(async (resolve, reject) => {
    try {
      if (!fs.existsSync(manifestPath)) {
        return reject(['The manifest to validate could not be found: ', manifestPath]);
      }

      const result = await officeAddinValidator.validateManifest(manifestPath);
      resolve(result);
    }
    catch (err) {
      return reject(err);
    }
  });
}

function getIdsAndManifests(application: string): Promise<Array<[string, string]>> {
  return new Promise(async (resolve, reject) => {
    try {
      const manifests = await getManifests(application);
      const ids = await getInfoForManifests(manifests);
      let idsAndManifests = [];

      for (let i = 0; i < manifests.length; i++) {
        idsAndManifests.push([ids[i], manifests[i]]);
      }

      resolve(idsAndManifests);
    } catch (err) {
      return reject(err);
    }
  });
}

function parseManifest(manifestPath: string): Promise<[string, string, string]> {
  return new Promise(async (resolve, reject) => {
    try {
      const parser = new xml2js.Parser();

      // Parse the manifest and get the id and version
      let manifestBuffer = await fs.readFile(manifestPath).catch((err) => {
        return reject(['Failed to read the manifest file: ', manifestPath]);
      });

      parser.parseString(manifestBuffer, (err, manifestXml) => {
        if (!manifestXml || typeof(manifestXml) !== 'object') {
          reject(['Failed to parse the manifest file: ', manifestPath]);
        }
        else if (!('OfficeApp' in manifestXml)) {
          return reject(['OfficeApp missing in manifest file: ', manifestPath]);
        }
        else if (!('$' in manifestXml['OfficeApp'] &&
            typeof(manifestXml['OfficeApp']['$'] === 'object') &&
            'xsi:type' in manifestXml['OfficeApp']['$'] &&
            typeof(manifestXml['OfficeApp']['$']['xsi:type'] === 'string'))) {
          return reject(['xsi:type missing in manifest file: ', manifestPath]);
        }
        else if (!('Id' in manifestXml['OfficeApp'] && manifestXml['OfficeApp']['Id'] instanceof Array)) {
          return reject(['Id missing in in manifest file: ', manifestPath]);
        }
        else if (!('Version' in manifestXml['OfficeApp'] && manifestXml['OfficeApp']['Version'] instanceof Array)) {
          return reject(['Version missing in in manifest file: ', manifestPath]);
        }

        const type = manifestXml['OfficeApp']['$']['xsi:type'];
        const id = manifestXml['OfficeApp']['Id'][0];
        const version = manifestXml['OfficeApp']['Version'][0];

        if (!isGuid(id)) {
          return reject(['Invalid Id ' + id + ' in manifest file: ', manifestPath]);
        }
        else if (type === 'MailApp') {
          return reject('The manifest specified an Outlook add-in. Outlook Add-ins are not supported by this tool');
        }
        else if (type !== 'ContentApp' && type !== 'TaskPaneApp') {
          return reject('The manifest must have xsi:type set to ContentApp or TaskPaneApp');
        }

        resolve ([type, id, version]);
      });
    }
    catch (err) {
      return reject(err);
    }
  });
}

function getInfoForManifests(manifestPaths: string[]): Promise<any> {
  return new Promise(async (resolve, reject) => {
    let ids = [];
    for (let manifestPath of manifestPaths) {
      try {
        const [type, id, version] = await parseManifest(manifestPath);
        ids.push(id);
      }
      catch (err) {
        console.log(err);
        ids.push(null);
      }
    }
    resolve(ids);
  });
}

function generateTemplateFile(application: string, type: string, id: string, version: string): Promise<any> {
  return new Promise(async (resolve, reject) => {
    try {
      if (Object.keys(applicationProperties).indexOf(application) < 0 ||
        Object.keys(applicationProperties[application]).indexOf(type) < 0) {
        return reject('The Add-in type ' + type + ' specified in the manifest is not supported for ' + application);
      }

      const defaultTemplateName = applicationProperties[application][type].templateName;
      const webExtensionPath = applicationProperties[application][type].webExtensionPath;
      let templatePath = path.join(process.cwd(), defaultTemplateName);

      let i = 0;
      while (fs.existsSync(templatePath)) {
        const [templateName, templateExtension] = defaultTemplateName.split('.');
        templatePath = path.join(process.cwd(), templateName + i + '.' + templateExtension);
        i++;
      }

      console.log(`Generating file ${templatePath}`);

      // Read the template
      const templateBuffer = await fs.readFile(path.join(__filename, '..', '..', 'templates', defaultTemplateName));
      const zip = await jszip.loadAsync(templateBuffer);

      // Replace the placeholder ID and version
      let webExtensionXml = await zip.file(webExtensionPath).async("string");
      webExtensionXml = webExtensionXml.replace(/00000000-0000-0000-0000-000000000000/g, id);
      webExtensionXml = webExtensionXml.replace(/1.0.0.0/g, version);
      zip.file(webExtensionPath, webExtensionXml);

      // Write the file
      zip.generateNodeStream({type: 'nodebuffer', streamFiles: true})
        .pipe(fs.createWriteStream(templatePath))
        .on('finish', () => {
          resolve(templatePath);
        });
    } catch (err) {
      return reject(err);
    }
  });
}
