val scala3Version = "3.0.1"

lazy val root = project
  .in(file("."))
  .settings(
    name := "finalibre-excel",
    version := "0.1.0",

    scalaVersion := scala3Version,

    libraryDependencies ++= List(
      "com.typesafe" % "config" % "1.4.1",
      "commons-io" % "commons-io" % "2.11.0",
      "org.apache.poi" % "poi-ooxml" % "5.0.0",
      "org.scalatest" %% "scalatest" % "3.2.9" % Test,
	  "ch.qos.logback" % "logback-classic" % "1.2.3"
    )
  )
