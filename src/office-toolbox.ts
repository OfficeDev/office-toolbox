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
import * as yeoman from 'yeoman-environment';

import * as util from './util';

function logRejection (err) {
  // When the error might contain personally identifiable information, only track the generic part.
  if (err instanceof Array && err.length) {
    util.appInsights.trackException(err[0]);
    for (let message of err) {
      console.log(chalk.red(message));
    }
  }
  else {
    util.appInsights.trackException(err[0]);
    console.log(chalk.red(err));
  }
}

// PROMPT FUNCTIONS //
async function promptForCommand () {
  const question = {
    name: 'command',
    type: 'list',
    message: 'What do you want to do?',
    choices: ['List registered developer manifests',
              'Sideload a manifest',
              'Remove a manifest',
              'Validate a manifest',
              'Generate a manifest']
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
      case 4:
        generate();
        break;
      default:
        commander.help();
    }
  });
}

async function promptForApplication () : Promise<string> {
  const question = {
    name: 'application',
    type: 'list',
    message: 'Which application are you targeting?',
    choices: ['Word', 'Excel', 'PowerPoint']
  };
  return inquirer.prompt(question).then((answer) => {
    return Promise.resolve(answer['application'].toLowerCase());
  });
}

async function checkAndPromptForPath (application: string, manifestPath: string) : Promise<string> {
  if (manifestPath) {
    return manifestPath;
  }
  else {
    console.log('The path must be specified for the manifest.');
    return promptForPathOrChoose().then((manifestSelectionMethod) => {
      if (manifestSelectionMethod === 'path') {
        return promptForManifestPath();
      }
      else if (manifestSelectionMethod === 'search') {
        return promptForManifestFromCurrentDirectory();
      }
      else if (manifestSelectionMethod === 'registered') {
        return promptForManifestFromListOfRegisteredManifests(application);
      }
      else {
        throw('An invalid method of specifying the manifest was selected.');
      }
    });
  }
}

async function promptForPathOrChoose () : Promise<string> {
  const question = {
    name: 'pathorchoose',
    type: 'list',
    message: 'Would you like to specify the path to a developer manifest or choose one that you have already registered?',
    choices: ['Specify the path to a developer manifest',
              'Choose a developer manifest from inside the current working directory',
              'Choose a registered developer manifest']
  };
  return inquirer.prompt(question).then((answer) => {
    switch (question.choices.indexOf(answer.pathorchoose)) {
      case 0:
        return Promise.resolve('path');
      case 1:
        return Promise.resolve('search');
      case 2:
        return Promise.resolve('registered');
    }
  });
}

async function promptForManifestFromListOfRegisteredManifests (application: string) : Promise<string> {
  if (!application && process.platform !== 'win32') {
    application = await promptForApplication();
  }
  const manifestPaths = await util.getManifests(application);
  return promptForManifestPathFromList(manifestPaths);
}

async function promptForManifestFromCurrentDirectory () : Promise<string> {
  const cwd = process.cwd();
  console.log('Searching for manifests in ' + cwd + '. This may take a while.');

  return getManifestsInDirectory(cwd, []).then((manifestPathsFoo) => {
    return promptForManifestPathFromList(manifestPathsFoo);
  });
}

async function promptForManifestPathFromList (manifestPaths: Array<string>) : Promise<string> {
  const question = {
    name: 'manifestPath',
    type: 'list',
    message: 'Choose a manifest:',
    choices: []
  };
  question.choices = [...manifestPaths];
  if (!question.choices.length) {
    return Promise.reject('There are no registered manifests to choose from.');
  }
  else {
    return inquirer.prompt(question).then((answers) => {
      return Promise.resolve(answers['manifestPath']);
    });
  }
}

// Recursively searches under the current directory for any files with extension .xml
function getManifestsInDirectory(directory: string, manifests: Array<string>) : Promise<Array<string>> {
  return new Promise (async (resolve, reject) => {
    fs.readdir(directory, (err, files) => {
      if (err) {
        resolve(manifests);
      }
      let promises = [];
      files.forEach(async (file) => {
        const fullPath = path.join(directory, file);
        if (fs.lstatSync(fullPath).isDirectory()) {
          promises.push(getManifestsInDirectory(fullPath, manifests));
        }
        else if (path.extname(file) === '.xml') {
          promises.push(fullPath);
        }
      });
      Promise.all(promises).then((values) => {
        resolve (Array.prototype.concat.apply(manifests, values));
      });
    });
  });
}

function promptForManifestPath () : Promise<string> {
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

    manifestPath = await checkAndPromptForPath(application, manifestPath);
    await util.sideload(application, manifestPath);
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

function generate() : Promise<any> {
  return new Promise((resolve, reject) => {
    try {
      util.appInsights.trackEvent('generate');
      const env = yeoman.createEnv();
      env.lookup(() => {
        env.run('office');
      });
    }
    catch (err) {
      return reject(err);
    }
  });
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
    sideload(options.application, options.manifest_path);
  });

commander
  .command('remove')
  .option('-a, --application <application>', 'The Office application. Word, PowerPoint, and Excel are currently supported. This parameter is ignored on Windows.')
  .option('-m, --manifest_path <manifest_path>', 'The path of the manifest file to remove.')
  .action(async (options) => {
    let application = (!options.application ? null : options.application.toLowerCase());
    remove(options.application, options.manifest_path);
  });

commander
  .command('validate')
  .option('-m, --manifest_path <manifest_path>', 'The path of the manifest file to validate.')
  .action(async (options) => {
    validate(options.manifest_path);
  });

commander
  .command('generate')
  .action(async (options) => {
    generate();
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
