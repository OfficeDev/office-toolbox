
import * as assert from "assert";
import * as fs from "fs";
import * as mocha from "mocha";
import * as officeToolbox from "../util";
import * as path from "path";
import * as util from "util";
const application = "excel";

describe("Add manifest", function () {
    describe(process.platform === "win32" ? "Add manifest to the registry" : "Add the manifest to the WEF folder", function () {
        this.timeout(0);
        let manifestFound = false;
        let manifests: string[];
        let testManifestPath = path.resolve(`${process.cwd()}/src/test/test-manifest.xml`);

        it(process.platform === "win32" ? "Manifest should have been added to the registry" : "Manifest should have been added to the WEF folder and given a unique name", async function () {
            await officeToolbox.addManifest(application, testManifestPath);

            if (process.platform === "win32") {
                manifests = await officeToolbox.getManifests(application);
            } else {
                testManifestPath = await officeToolbox.getSideloadManifestPath(testManifestPath, application);
                manifests = await officeToolbox.getManifestsFromSideloadingDirectory(application);

            }

            manifests.forEach(element => {
                if (element = testManifestPath) {
                    manifestFound = true;
                }
            });
            assert.equal(manifestFound, true);
        });
    });
});

describe("Remove manifest", function () {
    describe(process.platform === "win32" ? "Remove manifest from the registry" : "Remove uniquely-named manifest from the WEF folder", function () {
        this.timeout(0);
        let manifestFound = false;
        let manifests: string[];
        let testManifestPath = path.resolve(`${process.cwd()}/src/test/test-manifest.xml`);

        it(process.platform === "win32" ? "Manifest should have been removed from the registry" : "Manifest should have been removed from the WEF folder", async function () {
            await officeToolbox.removeManifest(application, testManifestPath, false /* manifestSelected */);

            if (process.platform === "win32") {
                manifests = await officeToolbox.getManifests(application);
            } else {
                testManifestPath = await officeToolbox.getSideloadManifestPath(testManifestPath, application);
                manifests = await officeToolbox.getManifestsFromSideloadingDirectory(application);

            }

            manifests.forEach(element => {
                if (element = testManifestPath) {
                    manifestFound = true;
                }
            });
            assert.equal(manifestFound, false);
        });
    });
});

// Tests are only relevant to Mac
if (process.platform == "darwin") {
    describe("Remove manifest with legacy name from sideloading directory (Mac only)", function () {
        describe("Remove the manifest from the WEF folder", function () {
            this.timeout(0);
            let manifestFound = false;
            let manifests: string[];
            let testManifestPath = path.resolve(`${process.cwd()}/src/test/test-manifest.xml`);

            it("Manifest should have been removed from the WEF folder", async function () {
                // Copy legacy manifest file name over to the WEF folder - this is neccessary because office-toolbox code no longer uses the legacy file name.                
                const copyFileAsync = util.promisify(fs.copyFile);
                const legacyManifestFilePath = `${officeToolbox.getSideloadingManifestDirectory(application)}/test-manifest.xml`
                await copyFileAsync(testManifestPath, legacyManifestFilePath);

                // Remove manifest and validate it's removed
                await officeToolbox.removeManifest(application, testManifestPath, false /* manifestSelected */);
                testManifestPath = await officeToolbox.getSideloadManifestPath(testManifestPath, application);
                manifests = await officeToolbox.getManifestsFromSideloadingDirectory(application);

                manifests.forEach(element => {
                    if (element = legacyManifestFilePath) {
                        manifestFound = true;
                    }
                });
                assert.equal(manifestFound, false);
            });
        });
    });

    describe("Remove manifest from WEF folder when manifest selected from prompt (Mac only)", function () {
        describe("Remove the manifest from the WEF folder", function () {
            this.timeout(0);
            let manifestFound = false;
            let manifests: string[];
            let testManifestPath = path.resolve(`${process.cwd()}/src/test/test-manifest.xml`);

            it("Manifest should have been removed from the WEF folder", async function () {
                // Copy legacy manifest file name over to the WEF folder - this is neccessary because office-toolbox code no longer uses the legacy file name.                
                const copyFileAsync = util.promisify(fs.copyFile);
                const legacyManifestFilePath = `${officeToolbox.getSideloadingManifestDirectory(application)}/test-manifest.xml`
                await copyFileAsync(testManifestPath, legacyManifestFilePath);

                // Remove manifest and validate it's removed
                await officeToolbox.removeManifest(application, legacyManifestFilePath, true /* manifestSelected */);
                testManifestPath = await officeToolbox.getSideloadManifestPath(testManifestPath, application);
                manifests = await officeToolbox.getManifestsFromSideloadingDirectory(application);

                manifests.forEach(element => {
                    if (element = legacyManifestFilePath) {
                        manifestFound = true;
                    }
                });
                assert.equal(manifestFound, false);
            });
        });
    });
}

