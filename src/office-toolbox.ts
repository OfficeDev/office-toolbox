#!/usr/bin/env node

/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

import * as chalk from 'chalk';
import * as commander from 'commander';
import * as fs from 'fs-extra';
import * as inquirer from 'inquirer';
import * as path from 'path';

import * as util from './util'; 

function logRejection(err) {
  let error: Error = undefined;

  if (err instanceof Array) {
     if (err.length) {
         error = (err[0] instanceof Error) ? err[0] : new Error(err[0]);
  
         for (const message of err) {
             console.log(chalk.default.red(message));
          }
      }
  }
  else {
      error = (err instanceof Error) ? err : new Error(err);
      console.log(chalk.default.red(err));  
  }
  
  util.appInsightsClient.trackException({ exception: error });
}

// PROMPT FUNCTIONS //
async function promptForCommand() {
  const question = {
    name: 'command',
    type: 'list',
    message: 'What do you want to do?',
    choices: ['List registered developer manifests',
              'Sideload a manifest',
              'Remove a manifest',
              'Validate a manifest']
  };
  await inquirer.prompt(question).then((answer) => {
    switch (question.choices.indexOf(answer.command)) {
      case 0:
        list();
        break;
      case 1:
        sideload(null, null);
        break;
      case 2:
        remove(null, null);
        break;
      case 3:
        validate(null);
        break;
      default:
        commander.help();
    }
  });
}

async function promptForApplication(): Promise<string> {
  const question = {
    name: 'application',
    type: 'list',
    message: 'Which application are you targeting?',
    choices: ['Word', 'Excel', 'PowerPoint', 'Outlook', 'OneNote', 'Project']
  };
  return inquirer.prompt(question).then((answer) => {
    return Promise.resolve(answer['application'].toLowerCase());
  });
}

async function checkAndPromptForPath(application: string, manifestPath: string): Promise<string> {
  if (manifestPath) { return manifestPath; }
  else {
    console.log('The path must be specified for the manifest.');

    return promptForPathOrChoose().then((manifestSelectionMethod) => {
      if (manifestSelectionMethod === 'path') { return promptForManifestPath(); }
      else if (manifestSelectionMethod === 'browse') { return promptForManifestFromCurrentDirectory(); }
      else if (manifestSelectionMethod === 'registered') { return promptForManifestFromListOfRegisteredManifests(application); }
      else { throw('An invalid method of specifying the manifest was selected.'); }
    });
  }
}

async function promptForPathOrChoose(): Promise<string> {
  const question = {
    name: 'pathorchoose',
    type: 'list',
    message: 'Would you like to specify the path to a developer manifest or choose one that you have already registered?',
    choices: ['Browse for a developer manifest from the current directory',
              'Specify the path to a developer manifest',
              'Choose a registered developer manifest']
  };
  return inquirer.prompt(question).then((answer) => {
    switch (question.choices.indexOf(answer.pathorchoose)) {
      case 0: return Promise.resolve('browse');
      case 1: return Promise.resolve('path');
      case 2: return Promise.resolve('registered');
    }
  });
}

async function promptForManifestFromListOfRegisteredManifests(application: string): Promise<string> {
  if (!application && process.platform !== 'win32') {
    application = await promptForApplication();
  }

  const manifestPaths = await util.getManifests(application);
  return promptForManifestPathFromList(manifestPaths, 'Choose a manifest:');
}

function promptForManifestFromCurrentDirectory(): Promise<string> {
  return new Promise (async (resolve, reject) => {
    const cwd = process.cwd();

    let manifestPath = cwd;
    while (fs.lstatSync(manifestPath).isDirectory()) {
      manifestPath = fs.realpathSync(manifestPath);
      const paths = await getItemsInDirectory(manifestPath);
      const choice = await promptForManifestPathFromList(paths, manifestPath);
      manifestPath = path.join(manifestPath, choice);
    }

    resolve(manifestPath);
  });
}

async function promptForManifestPathFromList(manifestPaths: string[], message: string): Promise<string> {
  const question = {
    name: 'manifestPath',
    type: 'list',
    message: message,
    choices: []
  };

  question.choices = [...manifestPaths];

  if (!question.choices.length) {
    return Promise.reject('There are no registered manifests to choose from.');
  }
  else {
    return inquirer.prompt(question).then(answers => {
      return Promise.resolve(answers['manifestPath']);
    });
  }
}

// Searches under the current directory for any files or extensions with extension .xml
function getItemsInDirectory(directory: string): string[] {
  let manifestPaths = [];
  let dirPaths = [".."];
  try {
    const files = fs.readdirSync(directory);
    files.forEach(async file => {
      const fullPath = path.join(directory, file);
      let stats;
      try {
        stats = fs.statSync(fullPath);
      }
      catch (e) {
        // Do nothing
      }

      if (stats && stats.isDirectory()) {
        dirPaths.push(file);
      }
      else if (path.extname(fullPath) === '.xml') {
        manifestPaths.push(file);
      }
    });
  }
  catch (err) {
    return null;
  }

  return [...manifestPaths, ...dirPaths];
}

function promptForManifestPath(): Promise<string> {
  const question = {
    name: 'manifestPath',
    type: 'input',
    message: 'Specify the path to the XML manifest file:',
  };

  return inquirer.prompt(question).then((answers) => {
    let manifestPath = answers['manifestPath'];

    if (manifestPath.charAt(0) === '"' && manifestPath.charAt(manifestPath.length - 1) === '"') {
      manifestPath = manifestPath.substr(1, manifestPath.length - 2);
    }

    return Promise.resolve(manifestPath);
  });
}

// TOP-LEVEL FUNCTIONS //
async function list() {
  try {
    const manifestInformation = await util.list();

    if (!manifestInformation || !manifestInformation.length) {
      console.log('No manifests were found.');
      return;
    }

    for (const [id, manifestPath, application] of (manifestInformation)) {
      let manifestString = (!application) ? '' : (application + ' ');
      manifestString += (!id) ? 'unknown                              ' : id + ' ';
      manifestString += manifestPath;

      console.log(manifestString);
    }
  }
  catch (err) {
    logRejection(err);
  }
}

async function sideload(application: string, manifestPath: string) {
  try {
    if (!application || Object.keys(util.applicationProperties).indexOf(application) < 0) {
      console.log('A valid application must be specified.');
      application = await promptForApplication();
    }

    const appProperties = util.applicationProperties[application];

    if (appProperties.canSideload) {
      manifestPath = await checkAndPromptForPath(application, manifestPath);
      await util.sideload(application, manifestPath);
      console.log(`For more information about how to sideload Office Add-ins, visit the following link: ${appProperties.documentationLink}`);
    }
    else {
      console.log(`Automatic sideloading is not available for this app, please follow the instructions in the following link: ${appProperties.documentationLink}`);
    }
    
  } catch (err) {
    logRejection(err);
  }
}

async function remove(application: string, manifestPath: string) {
  try {
    if (process.platform === 'win32') {
      application = null;
    }
    else if ((!application || Object.keys(util.applicationProperties).indexOf(application) < 0)) {
      console.log('A valid application must be specified.');
      application = await promptForApplication();
    }

    if (!manifestPath) {
      manifestPath = await promptForManifestFromListOfRegisteredManifests(application);
    }

    util.remove(application, manifestPath);
  } catch (err) {
    logRejection(err);
  }
}

async function validate(manifestPath: string) {
  try {
    manifestPath = await checkAndPromptForPath(null, manifestPath);
    await util.validate(manifestPath);
  }
  catch (err) {
    logRejection(err);
  }
}

// COMMANDER: Parse command-line input //
commander.on('--help', () => {
  console.log('  For help on a particular command, use:');
  console.log('');
  console.log('    office-toolbox [command] --help');
  console.log('');
});

commander
  .command('list')
  .action(() => {
    list();
  });

commander
  .command('sideload')
  .option('-a, --application <application>', 'The Office application. Word, Excel, and PowerPoint are currently supported.')
  .option('-m, --manifest_path <manifest_path>', 'The path of the manifest file to sideload and launch.')
  .action(async (options) => {
    let application = (!options.application ? null : options.application.toLowerCase());
    sideload(application, options.manifest_path);
  });

commander
  .command('remove')
  .option('-a, --application <application>', 'The Office application. Word, PowerPoint, and Excel are currently supported. This parameter is ignored on Windows.')
  .option('-m, --manifest_path <manifest_path>', 'The path of the manifest file to remove.')
  .action(async (options) => {
    let application = (!options.application ? null : options.application.toLowerCase());
    remove(application, options.manifest_path);
  });

commander
  .command('validate')
  .option('-m, --manifest_path <manifest_path>', 'The path of the manifest file to validate.')
  .action(async (options) => {
    validate(options.manifest_path);
  });

commander
  .command('*')
  .action(() => {
    commander.help();
  });

commander
  .parse(process.argv);

if (commander.args.length < 1) {
  promptForCommand();
}
