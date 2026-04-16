plugins {
    id("java-library")
    id("maven-publish")
    id("jacoco")
    id("me.champeau.jmh") version "0.6.8"
}

group = "io.github.scndry"
version = "1.0.0"
description = "Support for reading and writing Spreadsheet via Jackson abstractions."

val title = "Jackson dataformat: Spreadsheet"
val jacksonVersion = "2.21.2"
val poiVersion = "5.5.1"
val snapshots = version.toString().endsWith("SNAPSHOT")

repositories {
    mavenCentral()
}

dependencies {
    api("com.fasterxml.jackson.core:jackson-databind:$jacksonVersion")
    api("org.apache.poi:poi-ooxml:$poiVersion")
    implementation("com.fasterxml:aalto-xml:1.3.4")
    implementation("org.slf4j:slf4j-api:2.0.17")
    compileOnly("com.h2database:h2:2.2.224")
}

dependencies {
    testImplementation("org.assertj:assertj-core:3.27.7")
    testImplementation("org.junit.jupiter:junit-jupiter-api:6.0.3")
    testRuntimeOnly("ch.qos.logback:logback-classic:1.3.15")
    testRuntimeOnly("org.apache.logging.log4j:log4j-to-slf4j:2.24.3")
    testRuntimeOnly("org.junit.jupiter:junit-jupiter-engine:6.0.3")
}

dependencies {
    annotationProcessor("org.projectlombok:lombok:1.18.44")
    compileOnly("org.projectlombok:lombok:1.18.44")
    testAnnotationProcessor("org.projectlombok:lombok:1.18.44")
    testCompileOnly("org.projectlombok:lombok:1.18.44")
}

dependencies {
    jmh("com.alibaba:easyexcel:4.0.3") { exclude(group = "org.apache.poi") }
    jmh("org.dhatim:fastexcel:0.20.0")
    jmh("org.dhatim:fastexcel-reader:0.20.0")
    jmh("com.h2database:h2:2.2.224")
    jmh("com.fasterxml.woodstox:woodstox-core:7.1.1")
}

java {
    toolchain {
        languageVersion.set(JavaLanguageVersion.of(8))
    }
    withJavadocJar()
    withSourcesJar()
}

publishing {
    publications {
        create<MavenPublication>("maven") {
            from(components["java"])
            pom {
                name.set(title)
                description.set(project.description)
                url.set("https://github.com/scndry/jackson-dataformat-spreadsheet")
                licenses {
                    license {
                        name.set("The Apache License, Version 2.0")
                        url.set("https://www.apache.org/licenses/LICENSE-2.0.txt")
                    }
                }
                developers {
                    developer {
                        name.set("Seongman Yang")
                        email.set("scndryan@gmail.com")
                        url.set("https://scndry.github.io")
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
            name = "GitHubPackages"
            url = uri("https://maven.pkg.github.com/scndry/jackson-dataformat-spreadsheet")
            credentials {
                username = System.getenv("GITHUB_ACTOR") ?: findProperty("gpr.user") as String?
                password = System.getenv("GITHUB_TOKEN") ?: findProperty("gpr.key") as String?
            }
        }
    }
}

tasks.withType<Test> {
    useJUnitPlatform()
}

jmh {
    profilers.add("gc")
}

tasks.jacocoTestReport {
    reports {
        csv.required.set(true)
    }
}

tasks.withType<Javadoc> {
    val opts = options as CoreJavadocOptions
    opts.addBooleanOption("Xdoclint:-missing", true)
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
