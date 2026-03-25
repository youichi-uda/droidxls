pluginManagement {
    repositories {
        google()
        mavenCentral()
        gradlePluginPortal()
    }
}

@Suppress("UnstableApiUsage")
dependencyResolutionManagement {
    repositoriesMode.set(RepositoriesMode.FAIL_ON_PROJECT_REPOS)
    repositories {
        google()
        mavenCentral()
    }
}

rootProject.name = "droidxls"

include(":library")

// Composite build: depend on droidOffice-core source during development
includeBuild("../droidOffice-core") {
    dependencySubstitution {
        substitute(module("com.droidoffice:droidoffice-core")).using(project(":"))
    }
}
