# Publishing to Maven Central

This guide explains how to publish `html2excel-richtext` to Maven Central Repository.

## Prerequisites

### 1. Create Sonatype OSSRH Account

1. Create a JIRA account at https://issues.sonatype.org/
2. Create a "New Project" issue to request publishing rights for `io.github.fivefish130`
   - Project: Community Support - Open Source Project Repository Hosting (OSSRH)
   - Issue Type: New Project
   - Summary: Request to publish io.github.fivefish130
   - Group Id: io.github.fivefish130
   - Project URL: https://github.com/fivefish130/html2excel-richtext
   - SCM URL: https://github.com/fivefish130/html2excel-richtext.git
3. Wait for approval (usually within 2 business days)

**Important**: You must prove ownership of the GitHub account by:
- Creating a public GitHub repository named `OSSRH-xxxxx` (where xxxxx is your issue number), OR
- Adding the JIRA issue ID to your GitHub profile

### 2. Generate GPG Keys

```bash
# Generate GPG key pair
gpg --gen-key

# List your keys
gpg --list-keys

# Publish your public key to key servers
gpg --keyserver keyserver.ubuntu.com --send-keys YOUR_KEY_ID
gpg --keyserver keys.openpgp.org --send-keys YOUR_KEY_ID
gpg --keyserver pgp.mit.edu --send-keys YOUR_KEY_ID
```

### 3. Configure Maven Settings

Edit `~/.m2/settings.xml` (create if not exists):

```xml
<settings>
  <servers>
    <server>
      <id>ossrh</id>
      <username>YOUR_SONATYPE_USERNAME</username>
      <password>YOUR_SONATYPE_PASSWORD</password>
    </server>
  </servers>

  <profiles>
    <profile>
      <id>ossrh</id>
      <activation>
        <activeByDefault>true</activeByDefault>
      </activation>
      <properties>
        <gpg.executable>gpg</gpg.executable>
        <gpg.passphrase>YOUR_GPG_PASSPHRASE</gpg.passphrase>
      </properties>
    </profile>
  </profiles>
</settings>
```

**Security Note**: For production use, encrypt passwords using `mvn --encrypt-master-password` and `mvn --encrypt-password`.

## Publishing Steps

### 1. Verify Everything is Ready

```bash
# Run all tests
mvn clean test

# Verify build with all artifacts
mvn clean verify -Prelease
```

### 2. Deploy to Maven Central

```bash
# Deploy and release to Maven Central
mvn clean deploy -Prelease

# If autoReleaseAfterClose=false, you need to manually release:
# 1. Login to https://s01.oss.sonatype.org/
# 2. Go to "Staging Repositories"
# 3. Find your repository, "Close" it
# 4. After validation passes, "Release" it
```

### 3. Create GitHub Release

```bash
# Tag the release
git tag -a v1.0.0 -m "Release version 1.0.0"
git push origin v1.0.0

# Create GitHub release at:
# https://github.com/fivefish130/html2excel-richtext/releases/new
```

## Troubleshooting

### GPG Signing Fails

If you get "gpg: signing failed: Inappropriate ioctl for device":

```bash
export GPG_TTY=$(tty)
```

Add this to your `~/.bashrc` or `~/.zshrc` to make it permanent.

### Maven Central Sync Time

After release, it takes approximately:
- **10 minutes**: Available in staging repository
- **2 hours**: Available in Maven Central
- **4 hours**: Searchable on https://search.maven.org/

## Version Upgrade

For future releases:

1. Update version in all `pom.xml` files:
   ```bash
   mvn versions:set -DnewVersion=1.1.0
   mvn versions:commit
   ```

2. Update version in README.md and README_CN.md

3. Commit, tag, and deploy:
   ```bash
   git commit -am "chore: Bump version to 1.1.0"
   git tag -a v1.1.0 -m "Release version 1.1.0"
   git push origin main --tags
   mvn clean deploy -Prelease
   ```

## References

- [OSSRH Guide](https://central.sonatype.org/publish/publish-guide/)
- [Maven GPG Plugin](https://maven.apache.org/plugins/maven-gpg-plugin/)
- [Nexus Staging Plugin](https://github.com/sonatype/nexus-maven-plugins/tree/main/staging/maven-plugin)
