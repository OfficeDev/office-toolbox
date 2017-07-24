#!/usr/bin/env node

/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

import * as chalk from 'chalk';
import * as commander from 'commander';
import * as fs from 'fs-extra';
import * as inquirer from 'inquirer';
import * as util from './util';

function logRejection (err) {
  // When the error might contain personally identifiable information, only track the generic part.
  if (err instanceof Array && err.length > 0) {
    util.appInsights.trackException(err[0]);
    for (let message of err) {
      console.log(`${chalk.red(message)}`);
    }
  }
  else {
    util.appInsights.trackException(err[0]);
    console.log(`${chalk.red(err)}`);
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

function promptForApplication () : Promise<string> {
  return new Promise((resolve, reject) => {
    const question = {
      name: 'application',
      type: 'list',
      message: 'Which application are you targeting?',
      choices: ['Word', 'Excel', 'PowerPoint']
    };
    inquirer.prompt(question).then((answer) => {
      resolve(answer['application'].toLowerCase());
    });
  });
}

function checkAndPromptForPath (application: string, manifestPath: string) : Promise<string> {
  return new Promise(async (resolve, reject) => {
    try {
      if (manifestPath == null) {
        console.log('The path must be specified for the manifest.');
        const useManifestPath = (await promptForPathOrChoose() === 'path');
        manifestPath = useManifestPath ?
          await promptForManifestPath() :
          await promptForManifestPathFromList(application);
      }
      resolve (manifestPath);
    } catch (err) {
      return reject(err);
    }
  });
}

function promptForPathOrChoose () : Promise<string> {
  return new Promise((resolve, reject) => {
    const question = {
      name: 'pathorchoose',
      type: 'list',
      message: 'Would you like to specify the path to a developer manifest or choose one that you have already registered?',
      choices: ['Specify the path to a developer manifest',
                'Choose a registered developer manifest']
    };
    inquirer.prompt(question).then((answer) => {
      switch (question.choices.indexOf(answer.pathorchoose)) {
        case 0:
          resolve('path');
        case 1:
          resolve('choose');
      }
    });
  });
}

function promptForManifestPathFromList (application: string) : Promise<string> {
  return new Promise(async (resolve, reject) => {
    if (application == null && process.platform !== 'win32') {
      application = await promptForApplication();
    }
    const manifestPaths = await util.getManifests(application);
    const question = {
      name: 'manifestPath',
      type: 'list',
      message: 'Choose a manifest:',
      choices: []
    };
    for (const manifestPath of manifestPaths) {
      question.choices.push(manifestPath);
    }
    if (question.choices.length === 0) {
      return reject("There are no manifests registered to choose from.");
    }
    else {
      inquirer.prompt(question).then((answers) => {
        resolve(answers['manifestPath']);
      });
    }
  });
}

function promptForManifestPath () : Promise<string> {
  return new Promise((resolve, reject) => {
    const question = {
      name: 'manifestPath',
      type: 'input',
      message: 'Specify the path to the XML manifest file:',
    };
    inquirer.prompt(question).then((answers) => {
      let manifestPath = answers['manifestPath'];
      if (manifestPath.charAt(0) === '"' && manifestPath.charAt(manifestPath.length - 1) === '"') {
        manifestPath = manifestPath.substr(1, manifestPath.length - 2);
      }
      resolve(manifestPath);
    });
  });
}

// TOP-LEVEL FUNCTIONS //
async function list() {
  try {
    const manifestInformation = await util.list();
    if (manifestInformation == null || (manifestInformation).length === 0) {
      console.log('No manifests were found.');
      return;
    }
    for (const [id, manifestPath, application] of (manifestInformation)) {
      let manifestString = (application == null ) ? '' : (application + ' ');
      manifestString += (id == null) ? 'unknown                              ' : id + ' ';
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
    if (application == null || Object.keys(util.applicationProperties).indexOf(application) < 0) {
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
    else if ((application == null || Object.keys(util.applicationProperties).indexOf(application) < 0)) {
      console.log('A valid application must be specified.');
      application = await promptForApplication();
    }

    if (manifestPath == null) {
      manifestPath = await promptForManifestPathFromList(application);
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

function generate() {
  return new Promise((resolve, reject) => {
    try {
      util.appInsights.trackEvent('generate');
      const yeoman = require('yeoman-environment');
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
    let application = (options.application == null ? null : options.application.toLowerCase());
    sideload(options.application, options.manifest_path);
  });

commander
  .command('remove')
  .option('-a, --application <application>', 'The Office application. Word, PowerPoint, and Excel are currently supported. This parameter is ignored on Windows.')
  .option('-m, --manifest_path <manifest_path>', 'The path of the manifest file to remove.')
  .action(async (options) => {
    let application = (options.application == null ? null : options.application.toLowerCase());
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
