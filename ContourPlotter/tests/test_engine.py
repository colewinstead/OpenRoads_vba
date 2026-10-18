# SPDX-License-Identifier: GPL-2.0-or-later
"""Analytical and safety regressions; no ORD, QGIS, raster or company data."""
import json
import io
from contextlib import redirect_stderr
import math
from pathlib import Path
import tempfile
import unittest
import uuid
import numpy as np
from contour_plotter.tin_reader import read_tin, inspect, resolve_unit
from contour_plotter.contour_engine import extract, level_values
from contour_plotter.label_engine import place_labels, CollisionIndex
from contour_plotter.output_serializer import write_geometry, validate_geometry
from contour_plotter.cli import run, main

FIXTURE = Path(__file__).parent / "fixtures" / "openroads_minimal.xml"
SPATIAL = dict(axis_order="xy", vertical_unit="us-ft", model_unit="us-ft",
               crs="Verified synthetic local grid", vertical_datum="Synthetic local datum")


def xml(points, faces, extra="", units='linearUnit="meter" verticalLinearUnit="meter"'):
    p = "".join(f'<P id="{i+1}">{x} {y} {z}</P>' for i, (x, y, z) in enumerate(points))
    f = "".join(f"<F>{a+1} {b+1} {c+1}</F>" for a, b, c in faces)
    return f'<LandXML><Units><Metric {units}/></Units><Surfaces><Surface name="Test"><Definition surfType="TIN"><Pnts>{p}</Pnts><Faces>{f}</Faces></Definition></Surface>{extra}</Surfaces></LandXML>'


class ReaderTests(unittest.TestCase):
    def read(self, text, **kwargs):
        with tempfile.TemporaryDirectory() as tmp:
            p = Path(tmp) / "input.xml"
            p.write_text(text, encoding="utf-8")
            return read_tin(p, axis_order="xy", model_unit="m", crs="Local test grid",
                            vertical_datum="Local test datum", **kwargs)

    def test_reused_openroads_structure(self):
        p, f, m = read_tin(FIXTURE, **SPATIAL)
        self.assertEqual(p.tolist(), [[100, 200, 10], [110, 200, 11], [110, 210, 12]])
        self.assertEqual(f.tolist(), [[0, 1, 2]])
        self.assertEqual(m["vertical_unit"], "us-ft")

    def test_required_spatial_resolution(self):
        for missing in ("axis_order", "vertical_unit", "crs", "vertical_datum", "model_unit"):
            opts = dict(SPATIAL)
            del opts[missing]
            with self.subTest(missing=missing), self.assertRaises(ValueError):
                read_tin(FIXTURE, **opts)

    def test_axis_and_exact_foot_conversion(self):
        opts = dict(SPATIAL, axis_order="ne", model_unit="m")
        p, _, m = read_tin(FIXTURE, **opts)
        np.testing.assert_allclose(p[0], [200*1200/3937, 100*1200/3937, 10], rtol=0, atol=1e-12)
        self.assertEqual(m["axis_order"], "ne")
        self.assertNotEqual(1200/3937, .3048)

    def test_unit_contradiction(self):
        with self.assertRaisesRegex(ValueError, "conflicts"):
            read_tin(FIXTURE, **dict(SPATIAL, horizontal_unit="ft"))
        for declared in (None, "furlong"):
            with self.assertRaises(ValueError):
                resolve_unit(declared, None, "vertical")

    def test_multiple_and_duplicate_surface_selection(self):
        text = xml([(0, 0, 0), (1, 0, 1), (0, 1, 2)], [(0, 1, 2)])
        surf = text[text.index('<Surface name='):text.index('</Surfaces>')]
        text = text.replace("</Surfaces>", surf + "</Surfaces>")
        with self.assertRaisesRegex(ValueError, "Multiple"):
            self.read(text)
        with self.assertRaisesRegex(ValueError, "ambiguous"):
            self.read(text, surface_name="Test")
        self.assertEqual(self.read(text, surface_index=1)[2]["surface_index"], 1)

    def test_invisible_faces_and_unused_extreme_vertex(self):
        text = xml([(0, 0, 0), (1, 0, 1), (0, 1, 2), (10, 10, 1000)],
                   [(0, 1, 2), (1, 2, 3)])
        text = text.replace("<F>2 3 4", '<F i="1">2 3 4')
        p, f, meta = self.read(text)
        self.assertEqual(meta["invisible_faces"], 1)
        self.assertEqual(meta["zmax"], 2)
        self.assertEqual(len(f), 1)
        self.assertEqual(len(p), 3)

    def test_invalid_references_duplicate_ids_and_nonfinite(self):
        text = xml([(0, 0, 0), (1, 0, 1), (0, 1, 2)], [(0, 1, 2)])
        for bad in (text.replace("1 2 3</F>", "1 2 99</F>"),
                    text.replace('id="2"', 'id="1"'),
                    text.replace(">0 0 0</P>", ">nan 0 0</P>"),
                    text.replace("1 2 3</F>", "1 2 3 1</F>"),
                    text.replace("<F>", '<F i="maybe">')):
            with self.subTest(bad=bad), self.assertRaises(ValueError):
                self.read(bad)

    def test_degenerate_and_duplicate_triangles(self):
        with self.assertRaisesRegex(ValueError, "degenerate"):
            self.read(xml([(0, 0, 0), (1, 0, 1), (2, 0, 2)], [(0, 1, 2)]))
        with self.assertRaisesRegex(ValueError, "duplicate triangles"):
            self.read(xml([(0, 0, 0), (1, 0, 1), (0, 1, 2)], [(0, 1, 2), (2, 1, 0)]))

    def test_unit_conflicts_and_non_tin_rejected(self):
        source = xml([(0, 0, 0), (1, 0, 1), (0, 1, 2)], [(0, 1, 2)])
        with self.assertRaisesRegex(ValueError, "TIN Definition"):
            self.read(source.replace('surfType="TIN"', 'surfType="GRID"'))
        conflict = source.replace("</Units>", '<Imperial linearUnit="foot" verticalLinearUnit="foot"/></Units>')
        with self.assertRaisesRegex(ValueError, "conflicting"):
            self.read(conflict)

    def test_unsafe_xml_rejected(self):
        with self.assertRaisesRegex(ValueError, "DOCTYPE"):
            self.read('<!DOCTYPE LandXML [<!ENTITY e "attack">]><LandXML>&e;</LandXML>')

    def test_inspect_reports_unknown_vertical_without_inference(self):
        info = inspect(FIXTURE)
        self.assertIsNone(info["vertical_unit"])
        self.assertEqual(info["horizontal_unit"], "USSurveyFoot")


class ContourTests(unittest.TestCase):
    def test_known_single_triangle(self):
        p = np.array([[0., 0, 0], [10, 0, 10], [0, 10, 0]])
        c = extract(p, np.array([[0, 1, 2]]), minor=5, major=5)
        level = [v for v in c if v["elevation"] == 5]
        self.assertEqual(len(level), 1)
        self.assertEqual(set(map(tuple, level[0]["points"])), {(5., 0.), (5., 5.)})

    def test_shared_edge_connectivity_no_duplicates(self):
        p = np.array([[0., 0, 0], [10, 0, 10], [10, 10, 10], [0, 10, 0]])
        c = extract(p, np.array([[0, 1, 2], [0, 2, 3]]), minor=5, major=5)
        line = next(v for v in c if v["elevation"] == 5)
        self.assertEqual(len(line["points"]), 3)
        np.testing.assert_allclose(line["points"][:, 0], 5)
        self.assertEqual(sorted(line["points"][:, 1]), [0, 5, 10])

    def test_large_state_plane_coordinates(self):
        base = np.array([2_000_000., 1_000_000.])
        p = np.array([[0., 0, 0], [10, 0, 10], [0, 10, 0]])
        p[:, :2] += base
        c = extract(p, np.array([[0, 1, 2]]), minor=5, major=5)
        q = next(v for v in c if v["elevation"] == 5)["points"]
        self.assertEqual(set(map(tuple, q-base)), {(5., 0.), (5., 5.)})

    def test_closed_contour_pyramid(self):
        p = np.array([[0., 0, 0], [10, 0, 0], [10, 10, 0], [0, 10, 0], [5, 5, 10]])
        f = np.array([[0, 1, 4], [1, 2, 4], [2, 3, 4], [3, 0, 4]])
        c = extract(p, f, minor=5, major=5)
        q = next(v for v in c if v["elevation"] == 5)["points"]
        np.testing.assert_equal(q[0], q[-1])
        self.assertEqual(set(map(tuple, q)), {(2.5, 2.5), (7.5, 2.5), (7.5, 7.5), (2.5, 7.5)})

    def test_topological_hole_is_not_bridged(self):
        p = np.array([[x, y, x] for y in range(4) for x in range(4)], dtype=float)
        faces = []
        for y in range(3):
            for x in range(3):
                if (x, y) == (1, 1):
                    continue
                a = y*4+x
                faces += [[a, a+1, a+5], [a, a+5, a+4]]
        c = extract(p, np.array(faces), minor=.5, major=1)
        middle = [q["points"] for q in c if q["elevation"] == 1.5]
        self.assertEqual(len(middle), 2)
        self.assertFalse(any(q[:, 1].min() < 1 and q[:, 1].max() > 2 for q in middle))

    def test_flat_and_exact_vertex(self):
        f = np.array([[0, 1, 2]])
        self.assertEqual(extract(np.array([[0., 0, 1], [1, 0, 1], [0, 1, 1]]), f), [])
        p = np.array([[0., 0, 0], [10, 0, 10], [0, 10, 5]])
        q = next(c for c in extract(p, f, minor=5, major=5) if c["elevation"] == 5)["points"]
        self.assertEqual(set(map(tuple, q)), {(5., 0.), (0., 10.)})

    def test_decimal_classification_and_custom_base(self):
        levels = level_values(-.1, 1, .1, .5, .1)
        self.assertEqual([t for _, major, t in levels if major], ["0.1", "0.6"])
        with self.assertRaises(ValueError):
            level_values(0, 1, .3, .5)
        with self.assertRaises(ValueError):
            level_values(0, 1, 1e-6, 1)
        with self.assertRaises(ValueError):
            level_values(0, 1, float("nan"), 1)

    def test_cancel(self):
        with self.assertRaises(InterruptedError):
            extract(np.array([[0., 0, 0], [10, 0, 10], [0, 10, 0]]),
                    np.array([[0, 1, 2]]), cancel=lambda: True)


class LabelTests(unittest.TestCase):
    def line(self, ident=1, y=0, reverse=False):
        p = np.array([[0., y], [1000., y]])
        return dict(id=ident, elevation=100., major=True, text="100", points=p[::-1] if reverse else p)

    def test_spacing_scale_width_and_gaps(self):
        c = place_labels([self.line()], height=5, clearance=1, spacing=250, widths={"100": 12})[0]
        self.assertEqual(len(c["labels"]), 4)
        self.assertEqual(len(c["pieces"]), 5)
        length = sum(np.linalg.norm(np.diff(p, axis=0), axis=1).sum() for p in c["pieces"])
        self.assertAlmostEqual(length, 1000 - 4*14)
        self.assertEqual(c["labels"][0]["height"], .1*50)

    def test_readable_rotation(self):
        c = place_labels([self.line(reverse=True)], height=5)[0]
        self.assertTrue(all(abs(l["rotation"]) < 1e-12 for l in c["labels"]))

    def test_collisions_major_priority(self):
        contours = [dict(self.line(1), major=False), self.line(2, .1)]
        c = place_labels(contours, mode="both")
        self.assertEqual(len(c[0]["labels"]), 0)
        self.assertGreater(len(c[1]["labels"]), 0)

    def test_short_sharp_and_no_gaps(self):
        line = self.line()
        line["points"] = np.array([[0., 0], [1, 0]])
        self.assertEqual(place_labels([line])[0]["labels"], [])
        c = place_labels([self.line()], break_lines=False)[0]
        self.assertEqual(len(c["pieces"]), 1)
        np.testing.assert_equal(c["pieces"][0], self.line()["points"])

    def test_sharp_bends_are_not_labelled(self):
        contour = self.line()
        contour["points"] = np.array([[0., 0], [50, 0], [50, 50]])
        result = place_labels([contour], height=5, widths={"100": 20})[0]
        self.assertEqual(result["labels"], [])

    def test_measured_box_height_is_respected(self):
        contours = [self.line(1), self.line(2, 20)]
        a = place_labels(contours, box_height=5)
        b = place_labels(contours, box_height=30)
        self.assertGreater(len(a[1]["labels"]), len(b[1]["labels"]))

    def test_distinct_major_minor_heights_gaps_and_collision_boxes(self):
        major = dict(self.line(1), major=True, text="100")
        minor = dict(self.line(2, 30), major=False, text="101", elevation=101.)
        result = place_labels([major, minor], mode="both", height=6, minor_height=3,
                              widths={"100": 18, "101": 9}, box_height=6, minor_box_height=3)
        for contour, expected_height, width in zip(result, (6, 3), (18, 9)):
            self.assertTrue(contour["labels"])
            self.assertTrue(all(label["height"] == expected_height for label in contour["labels"]))
            self.assertTrue(all(abs((label["box"][3]-label["box"][1])-(expected_height+2)) < 1e-12 for label in contour["labels"]))
            length = sum(np.linalg.norm(np.diff(p, axis=0), axis=1).sum() for p in contour["pieces"])
            self.assertAlmostEqual(1000-length, len(contour["labels"])*(width+2))
        result = place_labels([major, minor], mode="both", height=3, minor_height=8)
        self.assertTrue(all(label["height"] == 8 for label in result[1]["labels"]))

    def test_conservative_collision_index(self):
        i = CollisionIndex(10)
        self.assertTrue(i.add_if_clear((0, 0, 20, 20)))
        self.assertFalse(i.add_if_clear((19, 19, 30, 30)))
        self.assertTrue(i.add_if_clear((21, 21, 30, 30)))


class OutputTests(unittest.TestCase):
    def test_roundtrip_and_corrupt_files(self):
        contours = place_labels([LabelTests().line()])
        meta = dict(vertical_unit="ft", model_unit="ft", bounds=[0., 0., 1000., 0.])
        rid = str(uuid.uuid4())
        with tempfile.TemporaryDirectory() as tmp:
            p = Path(tmp) / "geometry.cp2"
            stats = write_geometry(p, rid, contours, meta, 5)
            self.assertIn(b"\r\n", p.read_bytes())
            self.assertNotIn(b"\n", p.read_bytes().replace(b"\r\n", b""))
            counts = validate_geometry(p, rid)
            self.assertEqual(counts[:2], [stats["line_elements"], stats["text_elements"]])
            text = p.read_text()
            for broken in (text[:-10], text + "junk\n", text.replace("\t100\tM", "\tnan\tM", 1),
                           text.replace("\t125\t", "\t999999\t", 1)):
                if broken == text:
                    continue
                p.write_text(broken)
                with self.assertRaises(ValueError):
                    validate_geometry(p, rid)

    def test_long_line_chunks_and_3d_vertical_conversion(self):
        q = np.column_stack((np.arange(3000.), np.zeros(3000)))
        contours = [dict(id=1, elevation=10., major=True, text="10", points=q, pieces=[q], labels=[])]
        meta = dict(vertical_unit="ft", model_unit="m", bounds=[0., 0., 2999., 0.])
        with tempfile.TemporaryDirectory() as tmp:
            p = Path(tmp) / "geometry.cp2"
            stats = write_geometry(p, str(uuid.uuid4()), contours, meta, 5, dimension=3)
            self.assertEqual(stats["line_elements"], 3)
            self.assertAlmostEqual(float(p.read_text().splitlines()[2].split("\t")[3]), 3.048)
            validate_geometry(p)

    def test_source_cannot_be_overwritten_by_output(self):
        with tempfile.TemporaryDirectory() as tmp:
            source = Path(tmp) / "progress.json"
            source.write_bytes(FIXTURE.read_bytes())
            before = source.read_bytes()
            with self.assertRaisesRegex(ValueError, "outside the output"):
                run(dict(version=1, input=str(source), output_dir=tmp, spatial=SPATIAL))
            self.assertEqual(source.read_bytes(), before)

    def test_request_separate_sizes_and_analysis_protocol(self):
        with tempfile.TemporaryDirectory() as tmp:
            request = dict(version=1, action="analyze", input=str(FIXTURE), output_dir=tmp, spatial=SPATIAL,
                           settings=dict(minor=.5, major=1, label_mode="both", scale=1,
                                         paper_height_major=.2, paper_height_minor=.1, label_spacing=2))
            result = run(request)
            self.assertEqual((result["height"], result["minor_height"]), (.2, .1))
            fields = (Path(tmp)/"analysis.tsv").read_text(encoding="utf-16").strip().split("\t")
            self.assertEqual(len(fields), 14)
            self.assertEqual(tuple(map(float, fields[10:12])), (.2, .1))
            self.assertTrue(all(len(row.split("\t")) == 2 for row in (Path(tmp)/"levels.tsv").read_text().splitlines()))
            request["action"] = "generate"
            run(request)
            labels = [row.split("\t") for row in (Path(tmp)/"geometry.cp2").read_text().splitlines() if row.startswith("L\t")]
            self.assertEqual({row[3] for row in labels}, {"M", "m"})
            self.assertTrue(all(float(row[8]) == (.2 if row[3] == "M" else .1) for row in labels))
            request["settings"]["paper_height_minor"] = 0
            with self.assertRaisesRegex(ValueError, "paper_height_minor"):
                run(request)

    def test_class_specific_height_roundtrip_and_corruption(self):
        contours = [dict(LabelTests().line(1), major=True),
                    dict(LabelTests().line(2, 30), major=False, text="101", elevation=101.)]
        contours = place_labels(contours, mode="both", height=6, minor_height=3)
        metadata = dict(vertical_unit="ft", model_unit="ft", bounds=[0., 0., 1000., 30.])
        with tempfile.TemporaryDirectory() as tmp:
            p = Path(tmp)/"geometry.cp2"
            write_geometry(p, str(uuid.uuid4()), contours, metadata, 6, minor_height=3)
            validate_geometry(p)
            lines = p.read_text().splitlines()
            self.assertEqual(lines[0].split("\t")[7:9], ["6", "3"])
            for i, line in enumerate(lines):
                fields = line.split("\t")
                if fields[0] == "L" and fields[3] == "m":
                    fields[8] = "6"
                    lines[i] = "\t".join(fields)
                    break
            p.write_text("\n".join(lines)+"\n")
            with self.assertRaisesRegex(ValueError, "dimensions"):
                validate_geometry(p)

    def test_cli_error_message_and_source_protection(self):
        with tempfile.TemporaryDirectory() as tmp:
            root = Path(tmp)
            directory = root / "output"
            request_path = root / "request.json"
            request = dict(version=1, action="inspect", input=str(root/"missing.xml"), output_dir=str(directory))
            request_path.write_text(json.dumps(request), encoding="utf-8")
            with redirect_stderr(io.StringIO()):
                self.assertEqual(main(["--request", str(request_path)]), 2)
            error = json.loads((directory/"error.json").read_text())
            self.assertEqual((directory/"error.txt").read_text(encoding="utf-16"), error["error"])
            self.assertFalse((directory/"result.json").exists())
            source = directory/"error.txt"
            source.write_bytes(FIXTURE.read_bytes())
            before = source.read_bytes()
            request["input"] = str(source)
            request_path.write_text(json.dumps(request), encoding="utf-8")
            with redirect_stderr(io.StringIO()):
                self.assertEqual(main(["--request", str(request_path)]), 2)
            self.assertEqual(source.read_bytes(), before)

    def test_analyzed_source_identity_must_match(self):
        with tempfile.TemporaryDirectory() as tmp:
            req = dict(version=1, input=str(FIXTURE), output_dir=tmp, spatial=SPATIAL,
                       expected_source_sha256="0"*64, settings={})
            with self.assertRaisesRegex(ValueError, "changed after analysis"):
                run(req)
            self.assertFalse((Path(tmp)/"geometry.cp2").exists())

    def test_request_end_to_end_and_atomic_failure(self):
        with tempfile.TemporaryDirectory() as tmp:
            req = dict(version=1, input=str(FIXTURE), output_dir=tmp, spatial=SPATIAL,
                       settings=dict(minor=.5, major=1, label_mode="none"))
            result = run(req)
            self.assertTrue(result["ok"])
            p = Path(tmp) / "geometry.cp2"
            before = p.read_bytes()
            req["settings"]["minor"] = -1
            with self.assertRaises(ValueError):
                run(req)
            self.assertEqual(p.read_bytes(), before)
            self.assertEqual(json.loads((Path(tmp) / "result.json").read_text())["run_id"], result["run_id"])


if __name__ == "__main__":
    unittest.main()
