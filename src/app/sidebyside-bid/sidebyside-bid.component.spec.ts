import { async, ComponentFixture, TestBed } from '@angular/core/testing';

import { SidebysideBidComponent } from './sidebyside-bid.component';

describe('SidebysideBidComponent', () => {
  let component: SidebysideBidComponent;
  let fixture: ComponentFixture<SidebysideBidComponent>;

  beforeEach(async(() => {
    TestBed.configureTestingModule({
      declarations: [ SidebysideBidComponent ]
    })
    .compileComponents();
  }));

  beforeEach(() => {
    fixture = TestBed.createComponent(SidebysideBidComponent);
    component = fixture.componentInstance;
    fixture.detectChanges();
  });

  it('should create', () => {
    expect(component).toBeTruthy();
  });
});
