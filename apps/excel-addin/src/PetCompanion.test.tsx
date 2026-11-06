// @vitest-environment jsdom
import { useState } from 'react';
import { cleanup, fireEvent, render, screen } from '@testing-library/react';
import { afterEach, expect, it } from 'vitest';
import PetCompanion from './PetCompanion';

afterEach(cleanup);

it.each([false, true])('双击隐藏悬浮宠物，保留再次显示能力（busy=%s）', busy => {
  function Harness() {
    const [visible, setVisible] = useState(true);
    return <>{visible && <PetCompanion busy={busy} onHide={() => setVisible(false)} />}
      <button onClick={() => setVisible(true)}>显示助手</button></>;
  }
  render(<Harness />);
  const pet = screen.getByRole('button', {name:'查看技巧'});
  fireEvent.click(pet);
  expect(screen.getByRole('dialog', {name:'EB 小提示'})).toBeTruthy();
  fireEvent.click(pet);
  fireEvent.doubleClick(pet);
  expect(screen.queryByAltText('EB 助手')).toBeNull();
  expect(screen.queryByRole('dialog')).toBeNull();
  fireEvent.click(screen.getByRole('button', {name:'显示助手'}));
  expect(screen.getByRole('button', {name:'查看技巧'})).toBeTruthy();
});
