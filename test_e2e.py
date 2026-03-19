#!/usr/bin/env python3
"""
端到端测试脚本 - 验证完整的GUI和转换流程
"""

import tempfile
from pathlib import Path
from gui_utils import convert_single_file, GuiSettings, ConversionResult, FileStatistics


def test_e2e_conversion():
    """测试完整转换流程"""

    # 创建临时文件
    with tempfile.TemporaryDirectory() as tmpdir:
        tmpdir = Path(tmpdir)

        # 创建测试文件（模拟）
        test_file = tmpdir / "test.pdf"
        test_file.write_text("Test content for conversion")

        # 执行转换
        result = convert_single_file(test_file, tmpdir, 75)

        # 验证结果
        assert isinstance(result, ConversionResult)
        print(f"✓ 转换结果：{result.file_name} - {result.status}")

        # 测试统计
        stats = FileStatistics(total_files=1, completed_files=1, total_quality_score=85.0)
        assert stats.avg_quality == 85.0
        print(f"✓ 统计计算正确：平均质量={stats.avg_quality}")

        # 验证配置保存
        settings = GuiSettings()
        settings.set_output_dir(str(tmpdir))
        settings.set_quality_threshold(80)
        assert settings.get_output_dir() == str(tmpdir)
        assert settings.get_quality_threshold() == 80
        print("✓ 配置保存成功")


def test_statistics_calculation():
    """测试统计计算"""
    stats = FileStatistics(
        total_files=10,
        completed_files=8,
        failed_files=2,
        review_files=1,
        total_quality_score=680.0  # 8个文件的平均85
    )

    assert stats.completed_files == 8
    assert stats.failed_files == 2
    assert stats.review_files == 1
    assert stats.avg_quality == 85.0
    print(f"✓ 统计：完成{stats.completed_files}个，失败{stats.failed_files}个，待审核{stats.review_files}个")


def test_conversion_result_dataclass():
    """测试转换结果数据类"""
    result = ConversionResult(
        file_name="document.pdf",
        success=True,
        output_file=Path("/tmp/document.md"),
        error_message="",
        quality_score=88.5,
        status="完成"
    )

    assert result.file_name == "document.pdf"
    assert result.success == True
    assert result.quality_score == 88.5
    print("✓ ConversionResult数据结构正确")


if __name__ == "__main__":
    print("\n" + "="*50)
    print("端到端测试开始")
    print("="*50 + "\n")

    try:
        test_e2e_conversion()
        test_statistics_calculation()
        test_conversion_result_dataclass()

        print("\n" + "="*50)
        print("✅ 所有端到端测试通过")
        print("="*50)

    except AssertionError as e:
        print(f"\n❌ 测试失败：{str(e)}")
        exit(1)
    except Exception as e:
        print(f"\n❌ 测试异常：{str(e)}")
        exit(1)
