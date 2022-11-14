plugins {
    id("java-library")
    id("maven-publish")
}

group = "io.github.scndry"
version = "0.0.1-SNAPSHOT"
description = "Support for reading and writing Spreadsheet via Jackson abstractions."

val title = "Jackson dataformat: Spreadsheet"
val jacksonVersion = "2.14.0"
val poiVersion = "5.2.3"
val snapshots = version.toString().endsWith("SNAPSHOT")

repositories {
    mavenCentral()
}

dependencies {
    implementation("com.fasterxml.jackson.core:jackson-databind:$jacksonVersion")
    implementation("org.apache.poi:poi-ooxml:$poiVersion") {
        exclude("com.github.virtuald", "curvesapi")
        exclude("com.zaxxer", "SparseBitSet")
        exclude("commons-codec", "commons-codec")
        exclude("org.apache.commons", "commons-math3")
        exclude("org.apache.logging.log4j", "log4j-api")
    }
    implementation("org.slf4j:slf4j-api:1.7.36")
}

dependencies {
    testImplementation("com.navercorp.fixturemonkey:fixture-monkey:0.4.2")
    testImplementation("org.assertj:assertj-core:3.23.1")
    testImplementation("org.junit.jupiter:junit-jupiter-api:5.9.1")
    testRuntimeOnly("ch.qos.logback:logback-classic:1.3.4")
    testRuntimeOnly("org.apache.logging.log4j:log4j-to-slf4j:2.19.0")
    testRuntimeOnly("org.junit.jupiter:junit-jupiter-engine:5.9.1")
}

dependencies {
    annotationProcessor("org.projectlombok:lombok:1.18.24")
    compileOnly("org.projectlombok:lombok:1.18.24")
    testAnnotationProcessor("org.projectlombok:lombok:1.18.24")
    testCompileOnly("org.projectlombok:lombok:1.18.24")
}

java {
    withSourcesJar()
}

publishing {
    publications {
        create<MavenPublication>("maven") {
            from(components["java"])
            pom {
                name.set(title)
                description.set(project.description)
                licenses {
                    license {
                        name.set("The Apache License, Version 2.0")
                        url.set("https://www.apache.org/licenses/LICENSE-2.0.txt")
                    }
                }
                scm {
                    connection.set("scm:git:git://github.com/scndry/jackson-dataformat-spreadsheet.git")
                    developerConnection.set("scm:git:ssh://github.com/scndry/jackson-dataformat-spreadsheet.git")
                    url.set("https://github.com/scndry/jackson-dataformat-spreadsheet")
                }
            }
        }
    }
    repositories {
        maven {
            if ("repository" in properties) {
                name = properties["repository"] as String
                url = uri(properties[if (snapshots) "${name}Snapshots" else "${name}Releases"] as String)
                credentials(PasswordCredentials::class)
            } else {
                url = uri(layout.buildDirectory.dir(if (snapshots) "repository/snapshots" else "repository/releases"))
            }
        }
    }
}

tasks.withType<Test> {
    useJUnitPlatform()
}

tasks.withType<Jar> {
    manifest {
        attributes(
            "Implementation-Title" to title,
            "Implementation-Version" to archiveVersion,
            "Implementation-Vendor-Id" to project.group,
            "Specification-Title" to "jackson-databind",
            "Specification-Version" to jacksonVersion,
            "Specification-Vendor" to "FasterXML",
        )
    }
}
