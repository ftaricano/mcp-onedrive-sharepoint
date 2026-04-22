/**
 * Test suite for the new enhancement features
 * Validates cache, security, and improved functionality
 */

import {
  SecurityValidator,
  SecurePath,
  AuditLogger,
} from "./security-validator.js";
import {
  CacheManager,
  metadataCache,
  searchCache,
  driveCache,
} from "./cache-manager.js";

export class EnhancementTester {
  private testResults: Array<{
    test: string;
    passed: boolean;
    error?: string;
    duration?: number;
  }> = [];

  /**
   * Run all tests
   */
  async runAllTests(): Promise<void> {
    console.log("🧪 Running enhancement tests...\n");

    await this.testSecurityValidation();
    await this.testCacheManager();
    await this.testPathSecurity();
    await this.testAuditLogging();

    this.printResults();
  }

  /**
   * Test security validation features
   */
  private async testSecurityValidation(): Promise<void> {
    console.log("🔒 Testing Security Validation...");

    // Test path validation
    await this.runTest("Path Traversal Detection", () => {
      const result = SecurityValidator.validatePath("../../../etc/passwd");
      if (result.isValid)
        throw new Error("Should have detected path traversal");
    });

    await this.runTest("Valid Path Acceptance", () => {
      const result = SecurityValidator.validatePath(
        "documents/folder/file.txt",
      );
      if (!result.isValid) throw new Error("Should have accepted valid path");
    });

    // Test file name validation
    await this.runTest("Forbidden Extension Detection", () => {
      const result = SecurityValidator.validateFileName("malicious.exe");
      if (result.isValid)
        throw new Error("Should have rejected executable file");
    });

    await this.runTest("Reserved Name Detection", () => {
      const result = SecurityValidator.validateFileName("CON.txt");
      if (result.isValid) throw new Error("Should have rejected reserved name");
    });

    // Test search query validation
    await this.runTest("Script Injection Detection", () => {
      const result = SecurityValidator.validateSearchQuery(
        '<script>alert("xss")</script>',
      );
      if (result.isValid)
        throw new Error("Should have detected script injection");
    });

    // Test OData validation
    await this.runTest("OData Injection Detection", () => {
      const result = SecurityValidator.validateODataQuery({
        $filter: "name eq 'test' or '1'='1'",
        malicious: "drop table users",
      });
      if (result.isValid)
        throw new Error("Should have detected OData injection");
    });

    await this.runTest("Valid OData Query", () => {
      const result = SecurityValidator.validateODataQuery({
        $select: "name,size,lastModifiedDateTime",
        $top: 50,
      });
      if (!result.isValid)
        throw new Error("Should have accepted valid OData query");
    });
  }

  /**
   * Test cache manager functionality
   */
  private async testCacheManager(): Promise<void> {
    console.log("💾 Testing Cache Manager...");

    const testCache = new CacheManager<string>({
      maxSize: 3,
      defaultTTL: 1000,
      cleanupInterval: 0,
    });

    await this.runTest("Cache Set/Get", () => {
      testCache.set("key1", "value1");
      const result = testCache.get("key1");
      if (result !== "value1") throw new Error("Cache get failed");
    });

    await this.runTest("Cache Expiration", async () => {
      testCache.set("expiring", "value", 100); // 100ms TTL
      await new Promise((resolve) => setTimeout(resolve, 150));
      const result = testCache.get("expiring");
      if (result !== null) throw new Error("Cache should have expired");
    });

    await this.runTest("Cache LRU Eviction", () => {
      testCache.clear();
      testCache.set("key1", "value1");
      testCache.set("key2", "value2");
      testCache.set("key3", "value3");
      testCache.set("key4", "value4"); // Should evict key1

      if (testCache.get("key1") !== null)
        throw new Error("LRU eviction failed");
      if (testCache.get("key4") !== "value4")
        throw new Error("New key not stored");
    });

    await this.runTest("Cache Statistics", () => {
      testCache.clear();
      testCache.set("test", "value");
      testCache.get("test"); // hit
      testCache.get("missing"); // miss

      const stats = testCache.getStats();
      if (stats.hits !== 1 || stats.misses !== 1) {
        throw new Error("Cache statistics incorrect");
      }
    });

    testCache.destroy();
  }

  /**
   * Test path security utilities
   */
  private async testPathSecurity(): Promise<void> {
    console.log("🛡️ Testing Path Security...");

    await this.runTest("Secure Path Join", () => {
      const result = SecurePath.join("documents", "folder", "file.txt");
      if (!result.isValid || result.sanitized !== "documents/folder/file.txt") {
        throw new Error("Secure path join failed");
      }
    });

    await this.runTest("Path Traversal in Join", () => {
      const result = SecurePath.join("documents", "../../../etc", "passwd");
      if (result.isValid)
        throw new Error("Should have detected traversal in join");
    });

    await this.runTest("File Name Extraction", () => {
      const result = SecurePath.extractFileName("documents/folder/test.txt");
      if (!result.isValid || result.sanitized !== "test.txt") {
        throw new Error("File name extraction failed");
      }
    });

    await this.runTest("Parent Directory", () => {
      const result = SecurePath.getParentDir("documents/folder/file.txt");
      if (!result.isValid || result.sanitized !== "documents/folder") {
        throw new Error("Parent directory extraction failed");
      }
    });
  }

  /**
   * Test audit logging
   */
  private async testAuditLogging(): Promise<void> {
    console.log("📋 Testing Audit Logging...");

    await this.runTest("Basic Audit Log", () => {
      AuditLogger.clearLogs();
      AuditLogger.log(
        "test_operation",
        "test_user",
        "test_resource",
        "success",
      );

      const logs = AuditLogger.getRecentLogs(1);
      if (logs.length !== 1 || logs[0].operation !== "test_operation") {
        throw new Error("Audit logging failed");
      }
    });

    await this.runTest("Audit Log Sanitization", () => {
      AuditLogger.clearLogs();
      const maliciousInput = '<script>alert("xss")</script>';
      AuditLogger.log("test", "user", maliciousInput, "success");

      const logs = AuditLogger.getRecentLogs(1);
      if (logs[0].resource.includes("<script>")) {
        throw new Error("Audit log sanitization failed");
      }
    });

    await this.runTest("Log Entry Limit", () => {
      AuditLogger.clearLogs();

      // Add more than the limit to test truncation
      for (let i = 0; i < 1005; i++) {
        AuditLogger.log(`operation_${i}`, "user", "resource", "success");
      }

      const logs = AuditLogger.getRecentLogs(2000);
      if (logs.length > 1000) {
        throw new Error("Audit log limit not enforced");
      }
    });
  }

  /**
   * Test real cache performance
   */
  async testCachePerformance(): Promise<void> {
    console.log("⚡ Testing Cache Performance...");

    await this.runTest("Metadata Cache Performance", async () => {
      const testData = { id: "test", name: "Test File", size: 1024 };

      // Test cache performance
      const start = Date.now();
      for (let i = 0; i < 1000; i++) {
        metadataCache.set(`test_${i}`, testData);
      }
      const setTime = Date.now() - start;

      const start2 = Date.now();
      for (let i = 0; i < 1000; i++) {
        metadataCache.get(`test_${i}`);
      }
      const getTime = Date.now() - start2;

      console.log(
        `    📊 Cache Performance: Set=${setTime}ms, Get=${getTime}ms`,
      );

      if (setTime > 100 || getTime > 50) {
        throw new Error("Cache performance below expectations");
      }
    });

    await this.runTest("Search Cache Integration", () => {
      const query = "test search query";
      const results = [
        { id: "1", name: "result1" },
        { id: "2", name: "result2" },
      ];

      const cacheKey = searchCache.generateKey(query);
      searchCache.set(cacheKey, results);

      const cached = searchCache.get(cacheKey);
      if (!cached || cached.length !== 2) {
        throw new Error("Search cache integration failed");
      }
    });
  }

  /**
   * Helper method to run individual tests
   */
  private async runTest(
    testName: string,
    testFn: () => void | Promise<void>,
  ): Promise<void> {
    const start = Date.now();
    try {
      await testFn();
      const duration = Date.now() - start;
      this.testResults.push({ test: testName, passed: true, duration });
      console.log(`  ✅ ${testName} (${duration}ms)`);
    } catch (error) {
      const duration = Date.now() - start;
      this.testResults.push({
        test: testName,
        passed: false,
        error: error instanceof Error ? error.message : "Unknown error",
        duration,
      });
      console.log(
        `  ❌ ${testName} - ${error instanceof Error ? error.message : "Unknown error"} (${duration}ms)`,
      );
    }
  }

  /**
   * Print final test results
   */
  private printResults(): void {
    const passed = this.testResults.filter((r) => r.passed).length;
    const total = this.testResults.length;
    const totalTime = this.testResults.reduce(
      (sum, r) => sum + (r.duration || 0),
      0,
    );

    console.log("\n📊 Test Results Summary:");
    console.log(`  Total: ${total} tests`);
    console.log(`  Passed: ${passed} tests`);
    console.log(`  Failed: ${total - passed} tests`);
    console.log(`  Total Time: ${totalTime}ms`);
    console.log(`  Success Rate: ${((passed / total) * 100).toFixed(1)}%`);

    if (passed === total) {
      console.log("\n🎉 All tests passed! Enhancements are working correctly.");
    } else {
      console.log("\n⚠️  Some tests failed. Review the implementation.");

      const failed = this.testResults.filter((r) => !r.passed);
      console.log("\nFailed tests:");
      failed.forEach((test) => {
        console.log(`  • ${test.test}: ${test.error}`);
      });
    }
  }

  /**
   * Quick validation test for production
   */
  static async quickValidation(): Promise<boolean> {
    try {
      // Test critical security functions
      const pathTest = SecurityValidator.validatePath("../../../etc/passwd");
      const fileTest = SecurityValidator.validateFileName("test.exe");
      const searchTest = SecurityValidator.validateSearchQuery(
        "<script>alert(1)</script>",
      );

      if (pathTest.isValid || fileTest.isValid || searchTest.isValid) {
        console.error("❌ Security validation failed quick test");
        return false;
      }

      // Test cache functionality
      const testCache = new CacheManager({ maxSize: 2, defaultTTL: 1000 });
      testCache.set("test", "value");
      const cached = testCache.get("test");
      testCache.destroy();

      if (cached !== "value") {
        console.error("❌ Cache functionality failed quick test");
        return false;
      }

      console.log("✅ Quick validation passed");
      return true;
    } catch (error) {
      console.error("❌ Quick validation failed:", error);
      return false;
    }
  }
}

// Export for use in other modules
export { SecurityValidator, CacheManager, AuditLogger };

// Run tests if this file is executed directly
if (require.main === module) {
  const tester = new EnhancementTester();
  tester.runAllTests().catch(console.error);
}
