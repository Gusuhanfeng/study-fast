from __future__ import annotations

import copy
import importlib.util
from html.parser import HTMLParser
import json
from pathlib import Path
import shutil
import subprocess
import sys
import tempfile
import unittest

ROOT = Path(__file__).resolve().parents[1]
spec = importlib.util.spec_from_file_location('build_quiz', ROOT / 'scripts' / 'build_quiz.py')
quiz = importlib.util.module_from_spec(spec)
spec.loader.exec_module(quiz)


class QuizTests(unittest.TestCase):
    def setUp(self):
        self.bank = json.loads((ROOT / 'examples' / 'interactive-quiz.json').read_text(encoding='utf-8'))

    def test_normal_generation_and_refuse_overwrite(self):
        with tempfile.TemporaryDirectory() as directory:
            output = Path(directory) / 'quiz.html'
            args = [sys.executable, str(ROOT / 'scripts' / 'build_quiz.py'), '--config', str(ROOT / 'examples' / 'interactive-quiz.json'), '--output', str(output)]
            first = subprocess.run(args, capture_output=True)
            self.assertEqual(first.returncode, 0, first.stderr)
            original = output.read_bytes()
            second = subprocess.run(args, capture_output=True)
            self.assertNotEqual(second.returncode, 0)
            self.assertEqual(output.read_bytes(), original)

    def test_invalid_banks(self):
        for field, value in [('questions', []), ('title', ' '), ('bank_id', '../private')]:
            data = copy.deepcopy(self.bank); data[field] = value
            with self.subTest(field=field), self.assertRaises(ValueError):
                quiz.validate(data)
        self.bank['questions'].append(copy.deepcopy(self.bank['questions'][0]))
        with self.assertRaises(ValueError):
            quiz.validate(self.bank)

    def test_optional_teaching_fields_and_legacy_configuration(self):
        legacy = copy.deepcopy(self.bank)
        for question in legacy['questions']:
            question.pop('title', None)
            question.pop('learning', None)
        quiz.validate(legacy)
        for value in ({}, {'concept': ''}, {'unknown': 'text'}, {'example': []}, []):
            data = copy.deepcopy(self.bank)
            data['questions'][0]['learning'] = value
            with self.subTest(value=value), self.assertRaises(ValueError):
                quiz.validate(data)
        data = copy.deepcopy(self.bank)
        data['questions'][0]['title'] = ' '
        with self.assertRaises(ValueError):
            quiz.validate(data)

    def test_script_injection_stays_data(self):
        malicious = '</script><script>alert(1)</script><img src=x onerror=alert(2)>'
        self.bank['title'] = malicious
        self.bank['questions'][0]['answer'] = malicious
        self.bank['questions'][0]['title'] = malicious
        self.bank['questions'][0]['learning']['example'] = malicious
        rendered = quiz.render(self.bank)
        self.assertNotIn(malicious, rendered)
        self.assertIn('\\u003c/script\\u003e', rendered)
        self.assertEqual(rendered.count('<script>'), 1)
        self.bank['title'] = '__QUIZ_DATA__'
        self.assertIn('<title>__QUIZ_DATA__</title>', quiz.render(self.bank))

    def test_delivered_html_has_no_external_resources(self):
        class Resources(HTMLParser):
            external = []
            def handle_starttag(self, tag, attrs):
                for key, value in attrs:
                    if key in ('src', 'href', 'poster', 'srcset') and value and not value.startswith(('data:', '#')):
                        self.external.append((tag, key, value))
        parser = Resources()
        rendered = quiz.render(self.bank)
        parser.feed(rendered)
        self.assertEqual(parser.external, [])
        self.assertNotIn('@import', rendered)
        self.assertNotIn('fetch(', rendered)

    @unittest.skipUnless(shutil.which('node'), 'Node is needed for the JavaScript behavior checks')
    def test_runtime_sessions_storage_and_migration(self):
        # Execute the actual generated JS; a minimal DOM keeps this independent of a browser installation.
        harness = r'''
const fs=require('fs'),vm=require('vm'),assert=require('assert');
const html=fs.readFileSync(process.argv[1],'utf8');
const data=html.match(/<script id="quiz-data" type="application\/json">([\s\S]*?)<\/script>/)[1];
const script=html.match(/<script>([\s\S]*?)<\/script>/)[1];
const elements={};
function el(){return {textContent:'',value:'',hidden:false,dataset:{},children:[],setAttribute(){},addEventListener(){},appendChild(x){this.children.push(x)},replaceChildren(){this.children=[]}}}
const get=id=>elements[id]??=el();get('quiz-data').textContent=data;get('filter').value='all';get('order').value='sequential';
const saved={};const context=vm.createContext({document:{getElementById:get,createElement:el,querySelectorAll:()=>[],addEventListener(){}},localStorage:{getItem:k=>saved[k]||null,setItem:(k,v)=>saved[k]=v},confirm:()=>true});
vm.runInContext(script,context);const run=s=>vm.runInContext(s,context);
assert.equal(run('queue.length'),2);assert.equal(get('solution').hidden,true);
assert.equal(get('learningPanel').hidden,false);assert.equal(get('conceptBox').open,false);
assert.ok(get('concept').textContent.length>0);assert.equal(get('questionTitle').hidden,false);
get('conceptBox').open=true;run('next()');assert.equal(get('learningPanel').hidden,true);
assert.equal(get('questionTitle').hidden,true);run('previous()');assert.equal(get('conceptBox').open,false);
// Data in titles and teaching fields stays literal, never interpreted as markup.
run('bank.questions[0].learning.example="<img src=x onerror=alert(1)>";render()');
assert.equal(get('example').textContent,'<img src=x onerror=alert(1)>');
// A title-free legacy record still uses its original three-part fingerprint.
assert.equal(run('Object.keys(restore({version:1,bank_id:bank.bank_id,records:[{id:bank.questions[1].id,grade:"passed",fingerprint:JSON.stringify([bank.questions[1].question,bank.questions[1].answer,bank.questions[1].criteria||[]])}]})).length'),1);
run('grade("passed")');assert.equal(run('Object.keys(records).length'),0);
run('next()');assert.equal(run('stats(queue).unseen'),2); // skipping is not passing
run('previous();reveal();grade("review")');assert.equal(run('stats(queue).review'),1);
assert.equal(Object.keys(saved).length,1);
run('records=restore(JSON.parse(localStorage.getItem(key)))');assert.equal(run('stats(queue).review'),1);
run('start()');assert.equal(get('solution').hidden,true); // answer hidden on retest
get('draft').value='my working';run('next();previous()');assert.equal(get('draft').value,'my working');
get('search').value='不存在的题目';run('start()');assert.equal(run('queue.length'),0);
get('search').value='';run('start()');assert.equal(run('stats(queue,round).unseen'),2); // previous passes do not count this round
get('filter').value='review';run('start()');assert.equal(run('queue.length'),1);
run('reveal();grade("passed");next()');assert.equal(get('complete').hidden,false);
run('start()');assert.equal(run('queue.length'),0); // no stale passed item in review
get('filter').value='all';get('category').value='方法与应用';run('start()');assert.equal(run('queue.length'),1);
get('category').value='';get('order').value='random';run('start()');assert.equal(run('new Set(queue).size'),2);
run('bank.questions.reverse();records=restore(snapshot())');assert.equal(run('Object.keys(records).length'),1); // stable IDs survive reordering
run('bank.questions.find(q=>q.id==="concept-01").title="changed";records=restore(snapshot())');assert.equal(run('Object.keys(records).length'),0);
run('jump(queue.findIndex(id=>id==="concept-01"));reveal();grade("passed")');
run('bank.questions.find(q=>q.id==="concept-01").answer="revised";records=restore(snapshot())');assert.equal(run('Object.keys(records).length'),0);
assert.throws(()=>run('restore({version:1,bank_id:"other",records:[]})'));
assert.throws(()=>run('restore({version:1,bank_id:bank.bank_id,records:{}})'));
run('localStorage.setItem=()=>{throw Error("blocked")};save()');assert.ok(get('notice').textContent.length>0);
'''
        with tempfile.TemporaryDirectory() as directory:
            output = Path(directory) / 'quiz.html'
            output.write_text(quiz.render(self.bank), encoding='utf-8')
            result = subprocess.run(['node', '-e', harness, str(output)], capture_output=True, text=True, encoding='utf-8')
            self.assertEqual(result.returncode, 0, result.stderr)


if __name__ == '__main__':
    unittest.main()
